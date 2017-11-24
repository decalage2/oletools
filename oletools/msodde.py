#!/usr/bin/env python
"""
msodde.py

msodde is a script to parse MS Office documents
(e.g. Word, Excel), to detect and extract DDE links.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docx, .dotx, .docm, .dotm)
- Excel 2007+ (.xlsx, .xlsm)  (not .xlsb)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

msodde is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# msodde is copyright (c) 2017 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without modification,
# are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
# ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
# FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
# SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
# CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
# OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
# OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2017-10-18 v0.52 PL: - first version
# 2017-10-20       PL: - fixed issue #202 (handling empty xml tags)
# 2017-10-23       ES: - add check for fldSimple codes
# 2017-10-24       ES: - group tags and track begin/end tags to keep DDE strings together
# 2017-10-25       CH: - add json output
# 2017-10-25       CH: - parse doc
#                  PL: - added logging
# 2017-11-10       CH: - added field blacklist and corresponding cmd line args
# 2017-11-23       CH: - added support for xlsx files

__version__ = '0.52dev6'

#------------------------------------------------------------------------------
# TODO: field codes can be in headers/footers/comments - parse these
# TODO: generalize behaviour for xlsx: find all external links (maybe rename
#       command line flag for "blacklist" to "find all suspicious" or so)

#------------------------------------------------------------------------------
# REFERENCES:


#--- IMPORTS ------------------------------------------------------------------

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    import xml.etree.cElementTree as ET

import argparse
import zipfile
import os
import sys
import json
import logging
import re

# little hack to allow absolute imports even if oletools is not installed
# Copied from olevba.py
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty import olefile
import oletools.ooxml as ooxml
from oletools import xls_parser

# === PYTHON 2+3 SUPPORT ======================================================

if sys.version_info[0] >= 3:
    unichr = chr

# === CONSTANTS ==============================================================


NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NO_QUOTES = False
# XML tag for 'w:instrText'
TAG_W_INSTRTEXT = '{%s}instrText' % NS_WORD
TAG_W_FLDSIMPLE = '{%s}fldSimple' % NS_WORD
TAG_W_FLDCHAR = '{%s}fldChar' % NS_WORD
TAG_W_P = "{%s}p" % NS_WORD
TAG_W_R = "{%s}r" % NS_WORD
ATTR_W_INSTR = '{%s}instr' % NS_WORD
ATTR_W_FLDCHARTYPE = '{%s}fldCharType' % NS_WORD
LOCATIONS = ['word/document.xml','word/endnotes.xml','word/footnotes.xml','word/header1.xml','word/footer1.xml','word/header2.xml','word/footer2.xml','word/comments.xml']

# list of acceptable, harmless field instructions for blacklist field mode
# c.f. http://officeopenxml.com/WPfieldInstructions.php or the official
# standard ISO-29500-1:2016 / ECMA-376 paragraphs 17.16.4, 17.16.5, 17.16.23
# https://www.iso.org/standard/71691.html (neither mentions DDE[AUTO]).
# Format: (command, n_required_args, n_optional_args,
#          switches_with_args, switches_without_args, format_switches)
FIELD_BLACKLIST = (
    # date and time:
    ('CREATEDATE', 0, 0, '', 'hs',  'datetime'),
    ('DATE',       0, 0, '', 'hls', 'datetime'),
    ('EDITTIME',   0, 0, '', '',    'numeric'),
    ('PRINTDATE',  0, 0, '', 'hs',  'datetime'),
    ('SAVEDATE',   0, 0, '', 'hs',  'datetime'),
    ('TIME',       0, 0, '', '',    'datetime'),
    # exclude document automation (we hate the "auto" in "automation")
    # (COMPARE, DOCVARIABLE, GOTOBUTTON, IF, MACROBUTTON, PRINT)
    # document information
    ('AUTHOR',      0, 1, '', '',   'string'),
    ('COMMENTS',    0, 1, '', '',   'string'),
    ('DOCPROPERTY', 1, 0, '', '',   'string/numeric/datetime'),
    ('FILENAME',    0, 0, '', 'p',  'string'),
    ('FILESIZE',    0, 0, '', 'km', 'numeric'),
    ('KEYWORDS',    0, 1, '', '',   'string'),
    ('LASTSAVEDBY', 0, 0, '', '',   'string'),
    ('NUMCHARS',    0, 0, '', '',   'numeric'),
    ('NUMPAGES',    0, 0, '', '',   'numeric'),
    ('NUMWORDS',    0, 0, '', '',   'numeric'),
    ('SUBJECT',     0, 1, '', '',   'string'),
    ('TEMPLATE',    0, 0, '', 'p',  'string'),
    ('TITLE',       0, 1, '', '',   'string'),
    # equations and formulas
    # exlude '=' formulae because they have different syntax
    ('ADVANCE', 0, 0, 'dlruxy', '', ''),
    ('SYMBOL',  1, 0, 'fs', 'ahju', ''),
    # form fields
    ('FORMCHECKBOX', 0, 0, '', '', ''),
    ('FORMDROPDOWN', 0, 0, '', '', ''),
    ('FORMTEXT', 0, 0, '', '', ''),
    # index and tables
    ('INDEX', 0, 0, 'bcdefghklpsz', 'ry', ''),
    # exlude RD since that imports data from other files
    ('TA',  0, 0, 'clrs', 'bi', ''),
    ('TC',  1, 0, 'fl', 'n', ''),
    ('TOA', 0, 0, 'bcdegls', 'fhp', ''),
    ('TOC', 0, 0, 'abcdflnopst', 'huwxz', ''),
    ('XE',  1, 0, 'frty', 'bi', ''),
    # links and references
    # exclude AUTOTEXT and AUTOTEXTLIST since we do not like stuff with 'AUTO'
    ('BIBLIOGRAPHY', 0, 0, 'lfm', '', ''),
    ('CITATION', 1, 0, 'lfspvm', 'nty', ''),
    # exclude HYPERLINK since we are allergic to URLs
    # exclude INCLUDEPICTURE and INCLUDETEXT (other file or maybe even URL?)
    # exclude LINK and REF (could reference other files)
    ('NOTEREF', 1, 0, '', 'fhp', ''),
    ('PAGEREF', 1, 0, '', 'hp', ''),
    ('QUOTE', 1, 0, '', '', 'datetime'),
    ('STYLEREF', 1, 0, '', 'lnprtw', ''),
    # exclude all Mail Merge commands since they import data from other files
    # (ADDRESSBLOCK, ASK, COMPARE, DATABASE, FILLIN, GREETINGLINE, IF,
    #  MERGEFIELD, MERGEREC, MERGESEQ, NEXT, NEXTIF, SET, SKIPIF)
    # Numbering
    ('LISTNUM',      0, 1, 'ls', '', ''),
    ('PAGE',         0, 0, '', '', 'numeric'),
    ('REVNUM',       0, 0, '', '', ''),
    ('SECTION',      0, 0, '', '', 'numeric'),
    ('SECTIONPAGES', 0, 0, '', '', 'numeric'),
    ('SEQ',          1, 1, 'rs', 'chn', 'numeric'),
    # user information
    ('USERADDRESS', 0, 1, '', '', 'string'),
    ('USERINITIALS', 0, 1, '', '', 'string'),
    ('USERNAME', 0, 1, '', '', 'string'),
)

FIELD_DDE_REGEX = re.compile(r'^\s*dde(auto)?\s+', re.I)

FIELD_FILTER_DDE = 'only dde'
FIELD_FILTER_BLACKLIST = 'exclude blacklisted'
FIELD_FILTER_ALL = 'keep all'
FIELD_FILTER_DEFAULT = FIELD_FILTER_BLACKLIST


# banner to be printed at program start
BANNER = """msodde %s - http://decalage.info/python/oletools
THIS IS WORK IN PROGRESS - Check updates regularly!
Please report any issue at https://github.com/decalage2/oletools/issues
""" % __version__

BANNER_JSON = dict(type='meta', version=__version__, name='msodde',
                   link='http://decalage.info/python/oletools',
                   message='THIS IS WORK IN PROGRESS - Check updates regularly! '
                           'Please report any issue at '
                           'https://github.com/decalage2/oletools/issues')

# === LOGGING =================================================================

DEFAULT_LOG_LEVEL = "warning"  # Default log level
LOG_LEVELS = {
    'debug': logging.DEBUG,
    'info': logging.INFO,
    'warning': logging.WARNING,
    'error': logging.ERROR,
    'critical': logging.CRITICAL
}

class NullHandler(logging.Handler):
    """
    Log Handler without output, to avoid printing messages if logging is not
    configured by the main application.
    Python 2.7 has logging.NullHandler, but this is necessary for 2.6:
    see https://docs.python.org/2.6/library/logging.html#configuring-logging-for-a-library
    """
    def emit(self, record):
        pass

def get_logger(name, level=logging.CRITICAL+1):
    """
    Create a suitable logger object for this module.
    The goal is not to change settings of the root logger, to avoid getting
    other modules' logs on the screen.
    If a logger exists with same name, reuse it. (Else it would have duplicate
    handlers and messages would be doubled.)
    The level is set to CRITICAL+1 by default, to avoid any logging.
    """
    # First, test if there is already a logger with the same name, else it
    # will generate duplicate messages (due to duplicate handlers):
    if name in logging.Logger.manager.loggerDict:
        #NOTE: another less intrusive but more "hackish" solution would be to
        # use getLogger then test if its effective level is not default.
        logger = logging.getLogger(name)
        # make sure level is OK:
        logger.setLevel(level)
        return logger
    # get a new logger:
    logger = logging.getLogger(name)
    # only add a NullHandler for this logger, it is up to the application
    # to configure its own logging:
    logger.addHandler(NullHandler())
    logger.setLevel(level)
    return logger

# a global logger object used for debugging:
log = get_logger('msodde')


# === UNICODE IN PY2 =========================================================

def ensure_stdout_handles_unicode():
    """ Ensure stdout can handle unicode by wrapping it if necessary

    Required e.g. if output of this script is piped or redirected in a linux
    shell, since then sys.stdout.encoding is ascii and cannot handle
    print(unicode). In that case we need to find some compatible encoding and
    wrap sys.stdout into a encoder following (many thanks!)
    https://stackoverflow.com/a/1819009 or https://stackoverflow.com/a/20447935

    Can be undone by setting sys.stdout = sys.__stdout__
    """
    import codecs
    import locale

    # do not re-wrap
    if isinstance(sys.stdout, codecs.StreamWriter):
        return

    # try to find encoding for sys.stdout
    encoding = None
    try:
        encoding = sys.stdout.encoding  # variable encoding might not exist
    except Exception:
        pass

    if encoding not in (None, '', 'ascii'):
        return   # no need to wrap

    # try to find an encoding that can handle unicode
    try:
        encoding = locale.getpreferredencoding()
    except Exception:
        pass

    # fallback if still no encoding available
    if encoding in (None, '', 'ascii'):
        encoding = 'utf8'

    # logging is probably not initialized yet, but just in case
    log.debug('wrapping sys.stdout with encoder using {0}'.format(encoding))

    wrapper = codecs.getwriter(encoding)
    sys.stdout = wrapper(sys.stdout)


ensure_stdout_handles_unicode()   # e.g. for print(text) in main()


# === ARGUMENT PARSING =======================================================

class ArgParserWithBanner(argparse.ArgumentParser):
    """ Print banner before showing any error """
    def error(self, message):
        print(BANNER)
        super(ArgParserWithBanner, self).error(message)


def existing_file(filename):
    """ called by argument parser to see whether given file exists """
    if not os.path.exists(filename):
        raise argparse.ArgumentTypeError('File {0} does not exist.'
                                         .format(filename))
    return filename


def process_args(cmd_line_args=None):
    """ parse command line arguments (given ones or per default sys.argv) """
    parser = ArgParserWithBanner(description='A python tool to detect and extract DDE links in MS Office files')
    parser.add_argument("filepath", help="path of the file to be analyzed",
                        type=existing_file, metavar='FILE')
    parser.add_argument('-j', "--json", action='store_true',
                        help="Output in json format. Do not use with -ldebug")
    parser.add_argument("--nounquote", help="don't unquote values",action='store_true')
    parser.add_argument('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                        help="logging level debug/info/warning/error/critical (default=%(default)s)")
    filter_group = parser.add_argument_group(
         title='Filter which OpenXML field commands are returned',
         description='Only applies to OpenXML (e.g. docx), not to OLE (e.g. '
                     '.doc). These options are mutually exclusive, last option '
                     'found on command line overwrites earlier ones.')
    filter_group.add_argument('-d', '--dde-only', action='store_const',
                              dest='field_filter_mode', const=FIELD_FILTER_DDE,
                              help='Return only DDE and DDEAUTO fields')
    filter_group.add_argument('-f', '--filter', action='store_const',
                              dest='field_filter_mode', const=FIELD_FILTER_BLACKLIST,
                              help='Return all fields except harmless ones like PAGE')
    filter_group.add_argument('-a', '--all-fields', action='store_const',
                              dest='field_filter_mode', const=FIELD_FILTER_ALL,
                              help='Return all fields, irrespective of their contents')
    parser.set_defaults(field_filter_mode=FIELD_FILTER_DEFAULT)

    return parser.parse_args(cmd_line_args)


# === FUNCTIONS ==============================================================

# from [MS-DOC], section 2.8.25 (PlcFld):
# A field consists of two parts: field instructions and, optionally, a result. All fields MUST begin with
# Unicode character 0x0013 with sprmCFSpec applied with a value of 1. This is the field begin
# character. All fields MUST end with a Unicode character 0x0015 with sprmCFSpec applied with a value
# of 1. This is the field end character. If the field has a result, then there MUST be a Unicode character
# 0x0014 with sprmCFSpec applied with a value of 1 somewhere between the field begin character and
# the field end character. This is the field separator. The field result is the content between the field
# separator and the field end character. The field instructions are the content between the field begin
# character and the field separator, if one is present, or between the field begin character and the field
# end character if no separator is present. The field begin character, field end character, and field
# separator are collectively referred to as field characters.


def process_doc_field(data):
    """ check if field instructions start with DDE

    expects unicode input, returns unicode output (empty if not dde) """
    log.debug('processing field \'{0}\''.format(data))

    if data.lstrip().lower().startswith(u'dde'):
        #log.debug('--> is DDE!')
        return data
    elif data.lstrip().lower().startswith(u'\x00d\x00d\x00e\x00'):
        return data
    else:
        return u''


OLE_FIELD_START = 0x13
OLE_FIELD_SEP = 0x14
OLE_FIELD_END = 0x15
OLE_FIELD_MAX_SIZE = 1000   # max field size to analyze, rest is ignored


def process_doc_stream(stream):
    """ find dde links in single word ole stream

    since word ole file stream are subclasses of io.BytesIO, they are buffered,
    so reading char-wise is not that bad performanc-wise """

    have_start = False
    have_sep = False
    field_contents = None
    result_parts = []
    max_size_exceeded = False
    idx = -1
    while True:
        idx += 1
        char = stream.read(1)    # loop over every single byte
        if len(char) == 0:
            break
        else:
            char = ord(char)

        if char == OLE_FIELD_START:
            if have_start and max_size_exceeded:
                log.debug('big field was not a field after all')
            have_start = True
            have_sep = False
            max_size_exceeded = False
            field_contents = u''
            continue
        elif not have_start:
            continue

        # now we are after start char but not at end yet
        if char == OLE_FIELD_SEP:
            if have_sep:
                log.debug('unexpected field: has multiple separators!')
            have_sep = True
        elif char == OLE_FIELD_END:
            # have complete field now, process it
            new_result = process_doc_field(field_contents)
            if new_result:
                result_parts.append(new_result)

            # re-set variables for next field
            have_start = False
            have_sep = False
            field_contents = None
        elif not have_sep:
            # we are only interested in the part from start to separator
            # check that array does not get too long by accident
            if max_size_exceeded:
                pass
            elif len(field_contents) > OLE_FIELD_MAX_SIZE:
                log.debug('field exceeds max size of {0}. Ignore rest'
                          .format(OLE_FIELD_MAX_SIZE))
                max_size_exceeded = True

            # appending a raw byte to a unicode string here. Not clean but
            # all we do later is check for the ascii-sequence 'DDE' later...
            elif char == 0:        # may be a high-byte of a 2-byte codec
                field_contents += unichr(char)
            elif char in (10, 13):
                field_contents += u'\n'
            elif char < 32:
                field_contents += u'?'
            elif char < 128:
                field_contents += unichr(char)
            else:
                field_contents += u'?'

    if max_size_exceeded:
        log.debug('big field was not a field after all')

    log.debug('Checked {0} characters, found {1} fields'
              .format(idx, len(result_parts)))

    return result_parts


def process_doc(filepath):
    """
    find dde links in word ole (.doc/.dot) file

    like process_xml, returns a concatenated unicode string of dde links or
    empty if none were found. dde-links will still begin with the dde[auto] key
    word (possibly after some whitespace)
    """
    log.debug('process_doc')
    ole = olefile.OleFileIO(filepath, path_encoding=None)

    links = []
    for sid, direntry in enumerate(ole.direntries):
        is_orphan = direntry is None
        if is_orphan:
            # this direntry is not part of the tree --> unused or orphan
            direntry = ole._load_direntry(sid)
        is_stream = direntry.entry_type == olefile.STGTY_STREAM
        log.debug('direntry {:2d} {}: {}'
                  .format(sid, '[orphan]' if is_orphan else direntry.name,
                          'is stream of size {}'.format(direntry.size)
                          if is_stream else
                          'no stream ({})'
                          .format(direntry.entry_type)))
        if is_stream:
            new_parts = process_doc_stream(
                ole._open(direntry.isectStart, direntry.size))
            links.extend(new_parts)

    # mimic behaviour of process_docx: combine links to single text string
    return u'\n'.join(links)



def process_xls(filepath):
    """ find dde links in excel ole file """

    result = []
    for stream in xls_parser.XlsFile(filepath).get_streams():
        if not isinstance(stream, xls_parser.WorkbookStream):
            continue
        for record in stream.iter_records():
            if not isinstance(record, xls_parser.XlsRecordSupBook):
                continue
            if record.support_link_type in (
                    xls_parser.XlsRecordSupBook.LINK_TYPE_OLE_DDE,
                    xls_parser.XlsRecordSupBook.LINK_TYPE_EXTERNAL):
                result.append(record.virt_path)
    return u'\n'.join(result)


def process_docx(filepath, field_filter_mode=None):
    log.debug('process_docx')
    all_fields = []
    z = zipfile.ZipFile(filepath)
    for filepath in z.namelist():
        if filepath in LOCATIONS:
            data = z.read(filepath)
            fields = process_xml(data)
            if len(fields) > 0:
                #print ('DDE Links in %s:'%filepath)
                #for f in fields:
                #    print(f)
                all_fields.extend(fields)
    z.close()

    # apply field command filter
    log.debug('filtering with mode "{0}"'.format(field_filter_mode))
    if field_filter_mode in (FIELD_FILTER_ALL, None):
        clean_fields = all_fields
    elif field_filter_mode == FIELD_FILTER_DDE:
        clean_fields = [field for field in all_fields
                        if FIELD_DDE_REGEX.match(field)]
    elif field_filter_mode == FIELD_FILTER_BLACKLIST:
        # check if fields are acceptable and should not be returned
        clean_fields = [field for field in all_fields
                        if not field_is_blacklisted(field.strip())]
    else:
        raise ValueError('Unexpected field_filter_mode: "{0}"'
                         .format(field_filter_mode))

    return u'\n'.join(clean_fields)
    
def process_xml(data):
    # parse the XML data:
    root = ET.fromstring(data)
    fields = []
    ddetext = u''
    level = 0
    # find all the tags 'w:p':
    # parse each for begin and end tags, to group DDE strings
    # fldChar can be in either a w:r element, floating alone in the w:p or spread accross w:p tags
    # escape DDE if quoted etc
    # (each is a chunk of a DDE link)

    for subs in root.iter(TAG_W_P):
        elem = None
        for e in subs:
            #check if w:r and if it is parse children elements to pull out the first FLDCHAR or INSTRTEXT
            if e.tag == TAG_W_R:
                for child in e:
                    if child.tag == TAG_W_FLDCHAR or child.tag == TAG_W_INSTRTEXT:
                        elem = child
                        break
            else:
                elem = e
            #this should be an error condition
            if elem is None:
                continue
    
            #check if FLDCHARTYPE and whether "begin" or "end" tag
            if elem.attrib.get(ATTR_W_FLDCHARTYPE) is not None:
                if elem.attrib[ATTR_W_FLDCHARTYPE] == "begin":
                    level += 1    
                if elem.attrib[ATTR_W_FLDCHARTYPE] == "end":
                    level -= 1
                    if level == 0 or level == -1 : # edge-case where level becomes -1
                        fields.append(ddetext)
                        ddetext = u''
                        level = 0 # reset edge-case
        
            # concatenate the text of the field, if present:
            if elem.tag == TAG_W_INSTRTEXT and elem.text is not None:
                #expand field code if QUOTED
                ddetext += unquote(elem.text)

    for elem in root.iter(TAG_W_FLDSIMPLE):
        # concatenate the attribute of the field, if present:
        if elem.attrib is not None:
            fields.append(elem.attrib[ATTR_W_INSTR])

    return fields

def unquote(field): 
    if "QUOTE" not in field or NO_QUOTES:
        return field
    #split into components
    parts = field.strip().split(" ")
    ddestr = ""
    for p in parts[1:]:
        try: 
             ch = chr(int(p))
        except ValueError:
            ch = p
        ddestr += ch 
    return ddestr

# "static variables" for field_is_blacklisted:
FIELD_WORD_REGEX = re.compile(r'"[^"]*"|\S+')
FIELD_BLACKLIST_CMDS = tuple(field[0].lower() for field in FIELD_BLACKLIST)
FIELD_SWITCH_REGEX = re.compile(r'^\\[\w#*@]$')

def field_is_blacklisted(contents):
    """ Check if given field contents matches any in FIELD_BLACKLIST

    A complete parser of field contents would be really complicated, so this
    function has to make a trade-off. There may be valid constructs that this
    simple parser cannot comprehend. Most arguments are not tested for validity
    since that would make this test much more complicated. However, if this
    parser accepts some field contents, then office is very likely to not
    complain about it, either.
    """

    # split contents into "words", (e.g. 'bla' or '\s' or '"a b c"' or '""')
    words = FIELD_WORD_REGEX.findall(contents)
    if not words:
        return False

    # check if first word is one of the commands on our blacklist
    try:
        index = FIELD_BLACKLIST_CMDS.index(words[0].lower())
    except ValueError:    # first word is no blacklisted command
        return False
    log.debug('trying to match "{0}" to blacklist command {0}'
              .format(contents, FIELD_BLACKLIST[index]))
    _, nargs_required, nargs_optional, sw_with_arg, sw_solo, sw_format \
        = FIELD_BLACKLIST[index]

    # check number of args
    nargs = 0
    for word in words[1:]:
        if word[0] == '\\':  # note: words can never be empty, but can be '""'
            break
        nargs += 1
    if nargs < nargs_required:
        log.debug('too few args: found {0}, but need at least {1} in "{2}"'
                  .format(nargs, nargs_required, contents))
        return False
    elif nargs > nargs_required + nargs_optional:
        log.debug('too many args: found {0}, but need at most {1}+{2} in "{3}"'
                  .format(nargs, nargs_required, nargs_optional, contents))
        return False

    # check switches
    expect_arg = False
    arg_choices = []
    for word in words[1+nargs:]:
        if expect_arg:            # this is an argument for the last switch
            if arg_choices and (word not in arg_choices):
                log.debug('Found invalid switch argument "{0}" in "{1}"'
                          .format(word, contents))
                return False
            expect_arg = False
            arg_choices = []   # in general, do not enforce choices
            continue           # "no further questions, your honor"
        elif not FIELD_SWITCH_REGEX.match(word):
            log.debug('expected switch, found "{0}" in "{1}"'
                      .format(word, contents))
            return False
        # we want a switch and we got a valid one
        switch = word[1]

        if switch in sw_solo:
            pass
        elif switch in sw_with_arg:
            expect_arg = True     # next word is interpreted as arg, not switch
        elif switch == '#' and 'numeric' in sw_format:
            expect_arg = True     # next word is numeric format
        elif switch == '@' and 'datetime' in sw_format:
            expect_arg = True     # next word is date/time format
        elif switch == '*':
            expect_arg = True     # next word is format argument
            arg_choices += ['CHARFORMAT', 'MERGEFORMAT']  # always allowed
            if 'string' in sw_format:
                arg_choices += ['Caps', 'FirstCap', 'Lower', 'Upper']
            if 'numeric' in sw_format:
                arg_choices = []  # too many choices to list them here
        else:
            log.debug('unexpected switch {0} in "{1}"'.format(switch, contents))
            return False

    # if nothing went wrong sofar, the contents seems to match the blacklist
    return True


def process_xlsx(filepath, filed_filter_mode=None):
    """ process an OOXML excel file (e.g. .xlsx or .xlsb or .xlsm) """
    dde_links = []
    for subfile, elem, _ in ooxml.iter_xml(filepath):
        tag = elem.tag.lower()
        if tag == 'ddelink' or tag.endswith('}ddelink'):
            # we have found a dde link. Try to get more info about it
            link_info = ['DDE-Link']
            if 'ddeService' in elem.attrib:
                link_info.append(elem.attrib['ddeService'])
            if 'ddeTopic' in elem.attrib:
                link_info.append(elem.attrib['ddeTopic'])
            dde_links.append(' '.join(link_info))
    return u'\n'.join(dde_links)


def process_file(filepath, field_filter_mode=None):
    """ decides which of process_doc/x or process_xls/x to call """
    if olefile.isOleFile(filepath):
        if xls_parser.is_xls(filepath):
            return process_xls(filepath)
        else:
            return process_doc(filepath)
    try:
        doctype = ooxml.get_type(filepath)
        log.debug('Detected file type: {0}'.format(doctype))
        if doctype == ooxml.DOCTYPE_EXCEL:
            return process_xlsx(filepath, field_filter_mode)
        else:
            return process_docx(filepath, field_filter_mode)
    except Exception:
        return process_docx(filepath, field_filter_mode)


#=== MAIN =================================================================

def main(cmd_line_args=None):
    """ Main function, called if this file is called as a script

    Optional argument: command line arguments to be forwarded to ArgumentParser
    in process_args. Per default (cmd_line_args=None), sys.argv is used. Option
    mainly added for unit-testing
    """
    args = process_args(cmd_line_args)

    # Setup logging to the console:
    # here we use stdout instead of stderr by default, so that the output
    # can be redirected properly.
    logging.basicConfig(level=LOG_LEVELS[args.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)

    if args.json and args.loglevel.lower() == 'debug':
        log.warning('Debug log output will not be json-compatible!')

    if args.nounquote :
        global NO_QUOTES
        NO_QUOTES = True
        
    if args.json:
        jout = []
        jout.append(BANNER_JSON)
    else:
        # print banner with version
        print(BANNER)

    if not args.json:
        print('Opening file: %s' % args.filepath)

    text = ''
    return_code = 1
    try:
        text = process_file(args.filepath, args.field_filter_mode)
        return_code = 0
    except Exception as exc:
        if args.json:
            jout.append(dict(type='error', error=type(exc).__name__,
                             message=str(exc)))  # strange: str(exc) is enclosed in ""
        else:
            raise

    if args.json:
        for line in text.splitlines():
            if line.strip():
                jout.append(dict(type='dde-link', link=line.strip()))
        json.dump(jout, sys.stdout, check_circular=False, indent=4)
        print()   # add a newline after closing "]"
        return return_code  # required if we catch an exception in json-mode
    else:
        print ('DDE Links:')
        print(text)

    return return_code


if __name__ == '__main__':
    sys.exit(main())
