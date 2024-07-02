#!/usr/bin/env python
"""
msodde.py

msodde is a script to parse MS Office documents
(e.g. Word, Excel, RTF), to detect and extract DDE links.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docx, .dotx, .docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsx, .xlsm, .xlsb)
- RTF
- CSV (exported from / imported into Excel)
- XML (exported from Word 2003, Word 2007+, Excel 2003, (Excel 2007+?)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

msodde is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE =================================================================

# msodde is copyright (c) 2017-2019 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

# -- IMPORTS ------------------------------------------------------------------

from __future__ import print_function

import argparse
import os
import sys
import re
import csv

import olefile

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools import ooxml
from oletools import xls_parser
from oletools import rtfobj
from oletools.ppt_record_parser import is_ppt
from oletools import crypto
from oletools.common.io_encoding import ensure_stdout_handles_unicode
from oletools.common.log_helper import log_helper

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-10-18 v0.52 PL: - first version
# 2017-10-20       PL: - fixed issue #202 (handling empty xml tags)
# 2017-10-23       ES: - add check for fldSimple codes
# 2017-10-24       ES: - group tags and track begin/end tags to keep DDE
#                        strings together
# 2017-10-25       CH: - add json output
# 2017-10-25       CH: - parse doc
#                  PL: - added logging
# 2017-11-10       CH: - added field blacklist and corresponding cmd line args
# 2017-11-23       CH: - added support for xlsx files
# 2017-11-24       CH: - added support for xls files
# 2017-11-29       CH: - added support for xlsb files
# 2017-11-29       PL: - added support for RTF files (issue #223)
# 2017-12-07       CH: - ensure rtf file is closed
# 2018-01-05       CH: - add CSV
# 2018-01-11       PL: - fixed issue #242 (apply unquote to fldSimple tags)
# 2018-01-10       CH: - add single-xml files (Word 2003/2007+ / Excel 2003)
# 2018-03-21       CH: - added detection for various CSV formulas (issue #259)
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2018-10-25       CH: - detect encryption and raise error if detected
# 2019-03-25       CH: - added decryption of password-protected files
# 2019-07-17 v0.55 CH: - fixed issue #267, unicode error on Python 2


__version__ = '0.60.2'

# -----------------------------------------------------------------------------
# TODO: field codes can be in headers/footers/comments - parse these
# TODO: generalize behaviour for xlsx: find all external links (maybe rename
#       command line flag for "blacklist" to "find all suspicious" or so)
# TODO: Test with more interesting (real-world?) samples: xls, xlsx, xlsb, docx
# TODO: Think about finding all external "connections" of documents, not just
#       DDE-Links
# TODO: avoid reading complete rtf file data into memory

# -----------------------------------------------------------------------------
# REFERENCES:


# === PYTHON 2+3 SUPPORT ======================================================

if sys.version_info[0] >= 3:
    unichr = chr

# === CONSTANTS ==============================================================


NS_WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_WORD_2003 = 'http://schemas.microsoft.com/office/word/2003/wordml'
NO_QUOTES = False
# XML tag for 'w:instrText'
TAG_W_INSTRTEXT = ['{%s}instrText' % ns for ns in (NS_WORD, NS_WORD_2003)]
TAG_W_FLDSIMPLE = ['{%s}fldSimple' % ns for ns in (NS_WORD, NS_WORD_2003)]
TAG_W_FLDCHAR = ['{%s}fldChar' % ns for ns in (NS_WORD, NS_WORD_2003)]
TAG_W_P = ["{%s}p" % ns for ns in (NS_WORD, NS_WORD_2003)]
TAG_W_R = ["{%s}r" % ns for ns in (NS_WORD, NS_WORD_2003)]
ATTR_W_INSTR = ['{%s}instr' % ns for ns in (NS_WORD, NS_WORD_2003)]
ATTR_W_FLDCHARTYPE = ['{%s}fldCharType' % ns for ns in (NS_WORD, NS_WORD_2003)]
LOCATIONS = ('word/document.xml', 'word/endnotes.xml', 'word/footnotes.xml',
             'word/header1.xml', 'word/footer1.xml', 'word/header2.xml',
             'word/footer2.xml', 'word/comments.xml')

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
    # exlude '=' formulae because they have different syntax (and can be bad)
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

# filter modes
FIELD_FILTER_DDE = 'only dde'
FIELD_FILTER_BLACKLIST = 'exclude blacklisted'
FIELD_FILTER_ALL = 'keep all'
FIELD_FILTER_DEFAULT = FIELD_FILTER_BLACKLIST


# banner to be printed at program start
BANNER = """msodde %s - http://decalage.info/python/oletools
THIS IS WORK IN PROGRESS - Check updates regularly!
Please report any issue at https://github.com/decalage2/oletools/issues
""" % __version__

# === LOGGING =================================================================

DEFAULT_LOG_LEVEL = "warning"  # Default log level

# a global logger object used for debugging:
logger = log_helper.get_or_create_silent_logger('msodde')


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
    parser = ArgParserWithBanner(description='A python tool to detect and '
                                 'extract DDE links in MS Office files')
    parser.add_argument("filepath", help="path of the file to be analyzed",
                        type=existing_file, metavar='FILE')
    parser.add_argument('-j', "--json", action='store_true',
                        help="Output in json format. Do not use with -ldebug")
    parser.add_argument("--nounquote", help="don't unquote values",
                        action='store_true')
    parser.add_argument('-l', '--loglevel', dest="loglevel", action="store",
                        default=DEFAULT_LOG_LEVEL,
                        help="logging level debug/info/warning/error/critical "
                             "(default=%(default)s)")
    parser.add_argument("-p", "--password", type=str, action='append',
                        help='if encrypted office files are encountered, try '
                             'decryption with this password. May be repeated.')
    filter_group = parser.add_argument_group(
        title='Filter which OpenXML field commands are returned',
        description='Only applies to OpenXML (e.g. docx) and rtf, not to OLE '
                    '(e.g. .doc). These options are mutually exclusive, last '
                    'option found on command line overwrites earlier ones.')
    filter_group.add_argument('-d', '--dde-only', action='store_const',
                              dest='field_filter_mode', const=FIELD_FILTER_DDE,
                              help='Return only DDE and DDEAUTO fields')
    filter_group.add_argument('-f', '--filter', action='store_const',
                              dest='field_filter_mode',
                              const=FIELD_FILTER_BLACKLIST,
                              help='Return all fields except harmless ones')
    filter_group.add_argument('-a', '--all-fields', action='store_const',
                              dest='field_filter_mode', const=FIELD_FILTER_ALL,
                              help='Return all fields, irrespective of their '
                                   'contents')
    parser.set_defaults(field_filter_mode=FIELD_FILTER_DEFAULT)

    return parser.parse_args(cmd_line_args)


# === FUNCTIONS ==============================================================

# from [MS-DOC], section 2.8.25 (PlcFld):
# A field consists of two parts: field instructions and, optionally, a result.
# All fields MUST begin with Unicode character 0x0013 with sprmCFSpec applied
# with a value of 1. This is the field begin character. All fields MUST end
# with a Unicode character 0x0015 with sprmCFSpec applied with a value of 1.
# This is the field end character. If the field has a result, then there MUST
# be a Unicode character 0x0014 with sprmCFSpec applied with a value of 1
# somewhere between the field begin character and the field end character. This
# is the field separator. The field result is the content between the field
# separator and the field end character. The field instructions are the content
# between the field begin character and the field separator, if one is present,
# or between the field begin character and the field end character if no
# separator is present. The field begin character, field end character, and
# field separator are collectively referred to as field characters.


def process_doc_field(data):
    """ check if field instructions start with DDE

    expects unicode input, returns unicode output (empty if not dde) """
    logger.debug(u'processing field \'{0}\''.format(data))

    if data.lstrip().lower().startswith(u'dde'):
        return data
    if data.lstrip().lower().startswith(u'\x00d\x00d\x00e\x00'):
        return data
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
        if len(char) == 0:                   # pylint: disable=len-as-condition
            break
        else:
            char = ord(char)

        if char == OLE_FIELD_START:
            if have_start and max_size_exceeded:
                logger.debug('big field was not a field after all')
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
                logger.debug('unexpected field: has multiple separators!')
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
                logger.debug('field exceeds max size of {0}. Ignore rest'
                             .format(OLE_FIELD_MAX_SIZE))
                max_size_exceeded = True

            # appending a raw byte to a unicode string here. Not clean but
            # all we do later is check for the ascii-sequence 'DDE' later...
            elif char == 0:        # may be a high-byte of a 2-byte codec
                # pylint: disable-next=possibly-used-before-assignment
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
        logger.debug('big field was not a field after all')

    logger.debug('Checked {0} characters, found {1} fields'
                 .format(idx, len(result_parts)))

    return result_parts


def process_doc(ole):
    """
    find dde links in word ole (.doc/.dot) file

    Checks whether files is ppt and returns empty immediately in that case
    (ppt files cannot contain DDE-links to my knowledge)

    like process_xml, returns a concatenated unicode string of dde links or
    empty if none were found. dde-links will still begin with the dde[auto] key
    word (possibly after some whitespace)
    """
    logger.debug('process_doc')
    links = []
    for sid, direntry in enumerate(ole.direntries):
        is_orphan = direntry is None
        if is_orphan:
            # this direntry is not part of the tree --> unused or orphan
            direntry = ole._load_direntry(sid)
        is_stream = direntry.entry_type == olefile.STGTY_STREAM
        logger.debug('direntry {:2d} {}: {}'
                     .format(sid, '[orphan]' if is_orphan else direntry.name,
                             'is stream of size {}'.format(direntry.size)
                             if is_stream else
                             'no stream ({})'.format(direntry.entry_type)))
        if is_stream:
            new_parts = process_doc_stream(
                ole._open(direntry.isectStart, direntry.size))
            if new_parts:
                logger.debug("stream %r: %r" % (direntry.name, new_parts))
            links.extend(new_parts)

    # mimic behaviour of process_docx: combine links to single text string
    return u'\n'.join(links)


def process_xls(filepath):
    """ find dde links in excel ole file """

    result = []
    xls_file = None
    try:
        xls_file = xls_parser.XlsFile(filepath)
        for stream in xls_file.iter_streams():
            if not isinstance(stream, xls_parser.WorkbookStream):
                continue
            for record in stream.iter_records():
                if not isinstance(record, xls_parser.XlsRecordSupBook):
                    continue
                if record.support_link_type in (
                        xls_parser.XlsRecordSupBook.LINK_TYPE_OLE_DDE,
                        xls_parser.XlsRecordSupBook.LINK_TYPE_EXTERNAL):
                    result.append(record.virt_path.replace(u'\u0003', u' '))
        return u'\n'.join(result)
    finally:
        if xls_file is not None:
            xls_file.close()


def process_docx(filepath, field_filter_mode=None):
    """ find dde-links (and other fields) in Word 2007+ files """
    parser = ooxml.XmlParser(filepath)
    all_fields = []
    level = 0
    ddetext = u''
    for _, subs, depth in parser.iter_xml(tags=TAG_W_P + TAG_W_FLDSIMPLE):
        if depth == 0:   # at end of subfile:
            level = 0    # reset
        if subs.tag in TAG_W_FLDSIMPLE:
            # concatenate the attribute of the field, if present:
            attrib_instr = subs.attrib.get(ATTR_W_INSTR[0]) or \
                           subs.attrib.get(ATTR_W_INSTR[1])
            if attrib_instr is not None:
                all_fields.append(unquote(attrib_instr))
            continue

        # have a TAG_W_P
        for curr_elem in subs:
            # check if w:r; parse children to pull out first FLDCHAR/INSTRTEXT
            elem = None
            if curr_elem.tag in TAG_W_R:
                for child in curr_elem:
                    if child.tag in TAG_W_FLDCHAR or \
                            child.tag in TAG_W_INSTRTEXT:
                        elem = child
                        break
                if elem is None:
                    continue   # no fldchar or instrtext in this w:r
            else:
                elem = curr_elem
            if elem is None:
                raise ooxml.BadOOXML(filepath,
                                     'Got "None"-Element from iter_xml')

            # check if FLDCHARTYPE and whether "begin" or "end" tag
            attrib_type = elem.attrib.get(ATTR_W_FLDCHARTYPE[0]) or \
                          elem.attrib.get(ATTR_W_FLDCHARTYPE[1])
            if attrib_type is not None:
                if attrib_type == "begin":
                    level += 1
                if attrib_type == "end":
                    level -= 1
                    if level in (0, -1):  # edge-case; level gets -1
                        all_fields.append(ddetext)
                        ddetext = u''
                        level = 0  # reset edge-case

            # concatenate the text of the field, if present:
            if elem.tag in TAG_W_INSTRTEXT and elem.text is not None:
                # expand field code if QUOTED
                ddetext += unquote(elem.text)

    # apply field command filter
    logger.debug('filtering with mode "{0}"'.format(field_filter_mode))
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


def unquote(field):
    """TODO: document what exactly is happening here..."""
    if "QUOTE" not in field or NO_QUOTES:
        return field
    # split into components
    parts = field.strip().split(" ")
    ddestr = ""
    for part in parts[1:]:
        try:
            character = chr(int(part))
        except ValueError:
            character = part
        ddestr += character
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
    logger.debug(u'trying to match "{0}" to blacklist command {1}'
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
        logger.debug(u'too few args: found {0}, but need at least {1} in "{2}"'
                     .format(nargs, nargs_required, contents))
        return False
    if nargs > nargs_required + nargs_optional:
        logger.debug(u'too many args: found {0}, but need at most {1}+{2} in '
                     u'"{3}"'
                     .format(nargs, nargs_required, nargs_optional, contents))
        return False

    # check switches
    expect_arg = False
    arg_choices = []
    for word in words[1+nargs:]:
        if expect_arg:            # this is an argument for the last switch
            if arg_choices and (word not in arg_choices):
                logger.debug(u'Found invalid switch argument "{0}" in "{1}"'
                             .format(word, contents))
                return False
            expect_arg = False
            arg_choices = []   # in general, do not enforce choices
            continue           # "no further questions, your honor"
        elif not FIELD_SWITCH_REGEX.match(word):
            logger.debug(u'expected switch, found "{0}" in "{1}"'
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
            logger.debug(u'unexpected switch {0} in "{1}"'
                         .format(switch, contents))
            return False

    # if nothing went wrong sofar, the contents seems to match the blacklist
    return True


def process_xlsx(filepath):
    """ process an OOXML excel file (e.g. .xlsx or .xlsb or .xlsm) """
    dde_links = []
    parser = ooxml.XmlParser(filepath)
    for subfilename, elem, _ in parser.iter_xml():
        tag = elem.tag.lower()
        if tag == 'ddelink' or tag.endswith('}ddelink'):
            # we have found a dde link. Try to get more info about it
            link_info = []
            if 'ddeService' in elem.attrib:
                link_info.append(elem.attrib['ddeService'])
            if 'ddeTopic' in elem.attrib:
                link_info.append(elem.attrib['ddeTopic'])
            dde_links.append(u' '.join(link_info))
            logger.debug('Found tag "%s" in file %s: %s' % (tag, subfilename, repr(link_info)))

    # binary parts, e.g. contained in .xlsb
    for subfile, content_type, handle in parser.iter_non_xml():
        try:
            logger.info('Parsing non-xml subfile {0} with content type {1}'
                        .format(subfile, content_type))
            for record in xls_parser.parse_xlsb_part(handle, content_type,
                                                     subfile):
                logger.debug('{0}: {1}'.format(subfile, record))
                if isinstance(record, xls_parser.XlsbBeginSupBook) and \
                        record.link_type == \
                        xls_parser.XlsbBeginSupBook.LINK_TYPE_DDE:
                    dde_links.append(record.string1 + ' ' + record.string2)
        except Exception as exc:
            if content_type.startswith('application/vnd.ms-excel.') or \
               content_type.startswith('application/vnd.ms-office.'):  # pylint: disable=bad-indentation
                # should really be able to parse these either as xml or records
                log_func = logger.warning
            elif content_type.startswith('image/') or content_type == \
                    'application/vnd.openxmlformats-officedocument.' + \
                    'spreadsheetml.printerSettings':
                # understandable that these are not record-base
                log_func = logger.debug
            else:   # default
                log_func = logger.info
            log_func('Failed to parse {0} of content type {1} ("{2}")'
                     .format(subfile, content_type, str(exc)))
            # in any case: continue with next

    return u'\n'.join(dde_links)


class RtfFieldParser(rtfobj.RtfParser):
    """
    Specialized RTF parser to extract fields such as DDEAUTO
    """

    def __init__(self, data):
        super(RtfFieldParser, self).__init__(data)
        # list of RtfObjects found
        self.fields = []

    def open_destination(self, destination):
        if destination.cword == b'fldinst':
            logger.debug('*** Start field data at index %Xh'
                         % destination.start)

    def close_destination(self, destination):
        if destination.cword == b'fldinst':
            logger.debug('*** Close field data at index %Xh' % self.index)
            logger.debug('Field text: %r' % destination.data)
            # remove extra spaces and newline chars:
            field_clean = destination.data.translate(None, b'\r\n').strip()
            logger.debug('Cleaned Field text: %r' % field_clean)
            self.fields.append(field_clean)

    def control_symbol(self, matchobject):
        # required to handle control symbols such as '\\'
        # inject the symbol as-is in the text:
        # TODO: handle special symbols properly
        self.current_destination.data += matchobject.group()[1:2]


RTF_START = b'\x7b\x5c\x72\x74'   # == b'{\rt' but does not mess up auto-indent


def process_rtf(file_handle, field_filter_mode=None):
    """ find dde links or other fields in rtf file """
    all_fields = []
    data = RTF_START + file_handle.read()   # read complete file into memory!
    file_handle.close()
    rtfparser = RtfFieldParser(data)
    rtfparser.parse()
    all_fields = [field.decode('ascii') for field in rtfparser.fields]
    # apply field command filter
    logger.debug('found {1} fields, filtering with mode "{0}"'
                 .format(field_filter_mode, len(all_fields)))
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


# threshold when to consider a csv file "small"; also used as sniffing size
CSV_SMALL_THRESH = 1024

# format of dde link: program-name | arguments ! unimportant
# can be enclosed in "", prefixed with + or = or - or cmds like @SUM(...)
CSV_DDE_FORMAT = re.compile(r'\s*"?[=+-@](.+)\|(.+)!(.*)\s*')

# allowed delimiters (python sniffer would use nearly any char). Taken from
# https://data-gov.tw.rpi.edu/wiki/CSV_files_use_delimiters_other_than_commas
CSV_DELIMITERS = ',\t ;|^'


def process_csv(filepath):
    """ find dde in csv text

    finds text parts like =cmd|'/k ..\\..\\..\\Windows\\System32\\calc.exe'! or
    =MSEXCEL|'\\..\\..\\..\\Windows\\System32\\regsvr32 [...]

    Hoping here that the :py:class:`csv.Sniffer` determines quote and delimiter
    chars the same way that excel does. Tested to some extend in unittests.

    This can only find DDE-links, no other "suspicious" constructs (yet).

    Cannot deal with unicode files yet (need more than just use uopen()).
    """
    results = []
    if sys.version_info.major <= 2:
        open_arg = dict(mode='rb')
    else:
        open_arg = dict(newline='')
    with open(filepath, **open_arg) as file_handle:
        # TODO: here we should not assume this is a file on disk, filepath can be a file object
        results, dialect = process_csv_dialect(file_handle, CSV_DELIMITERS)
        is_small = file_handle.tell() < CSV_SMALL_THRESH

        if is_small and not results:
            # easy to mis-sniff small files. Try different delimiters
            logger.debug('small file, no results; try all delimiters')
            file_handle.seek(0)
            other_delim = CSV_DELIMITERS.replace(dialect.delimiter, '')
            for delim in other_delim:
                try:
                    file_handle.seek(0)
                    results, _ = process_csv_dialect(file_handle, delim)
                except csv.Error:   # e.g. sniffing fails
                    logger.debug('failed to csv-parse with delimiter {0!r}'
                                 .format(delim))

        if is_small and not results:
            # try whole file as single cell, since sniffing fails in this case
            logger.debug('last attempt: take whole file as single unquoted '
                         'cell')
            file_handle.seek(0)
            match = CSV_DDE_FORMAT.match(file_handle.read(CSV_SMALL_THRESH))
            if match:
                results.append(u' '.join(match.groups()[:2]))

    return u'\n'.join(results)


def process_csv_dialect(file_handle, delimiters):
    """ helper for process_csv: process with a specific csv dialect """
    # determine dialect = delimiter chars, quote chars, ...
    dialect = csv.Sniffer().sniff(file_handle.read(CSV_SMALL_THRESH),
                                  delimiters=delimiters)
    dialect.strict = False     # microsoft is never strict
    logger.debug('sniffed csv dialect with delimiter {0!r} '
                 'and quote char {1!r}'
                 .format(dialect.delimiter, dialect.quotechar))

    # rewind file handle to start
    file_handle.seek(0)

    # loop over all csv rows and columns
    results = []
    reader = csv.reader(file_handle, dialect)
    for row in reader:
        for cell in row:
            # check if cell matches
            match = CSV_DDE_FORMAT.match(cell)
            if match:
                results.append(u' '.join(match.groups()[:2]))
    return results, dialect


#: format of dde formula in excel xml files
XML_DDE_FORMAT = CSV_DDE_FORMAT


def process_excel_xml(filepath):
    """ find dde links in xml files created with excel 2003 or excel 2007+

    TODO: did not manage to create dde-link in the 2007+-xml-format. Find out
          whether this is possible at all. If so, extend this function
    """
    dde_links = []
    parser = ooxml.XmlParser(filepath)
    for _, elem, _ in parser.iter_xml():
        tag = elem.tag.lower()
        if tag != 'cell' and not tag.endswith('}cell'):
            continue   # we are only interested in cells
        formula = None
        for key in elem.keys():
            if key.lower() == 'formula' or key.lower().endswith('}formula'):
                formula = elem.get(key)
                break
        if formula is None:
            continue
        logger.debug(u'found cell with formula {0}'.format(formula))
        match = re.match(XML_DDE_FORMAT, formula)
        if match:
            dde_links.append(u' '.join(match.groups()[:2]))
    return u'\n'.join(dde_links)


def process_file(filepath, field_filter_mode=None):
    """ decides which of the process_* functions to call """
    if olefile.isOleFile(filepath):
        logger.debug('Is OLE. Checking streams to see whether this is xls')
        if xls_parser.is_xls(filepath):
            logger.debug('Process file as excel 2003 (xls)')
            return process_xls(filepath)
        if is_ppt(filepath):
            logger.debug('is ppt - cannot have DDE')
            return u''
        logger.debug('Process file as word 2003 (doc)')
        with olefile.OleFileIO(filepath, path_encoding=None) as ole:
            return process_doc(ole)

    with open(filepath, 'rb') as file_handle:
        # TODO: here we should not assume this is a file on disk, filepath can be a file object
        if file_handle.read(4) == RTF_START:
            logger.debug('Process file as rtf')
            return process_rtf(file_handle, field_filter_mode)

    try:
        doctype = ooxml.get_type(filepath)
        logger.debug('Detected file type: {0}'.format(doctype))
    except Exception as exc:
        logger.debug('Exception trying to xml-parse file: {0}'.format(exc))
        doctype = None

    if doctype == ooxml.DOCTYPE_EXCEL:
        logger.debug('Process file as excel 2007+ (xlsx)')
        return process_xlsx(filepath)
    if doctype in (ooxml.DOCTYPE_EXCEL_XML, ooxml.DOCTYPE_EXCEL_XML2003):
        logger.debug('Process file as xml from excel 2003/2007+')
        return process_excel_xml(filepath)
    if doctype in (ooxml.DOCTYPE_WORD_XML, ooxml.DOCTYPE_WORD_XML2003):
        logger.debug('Process file as xml from word 2003/2007+')
        return process_docx(filepath)
    if doctype is None:
        logger.debug('Process file as csv')
        return process_csv(filepath)
    # could be docx; if not: this is the old default code path
    logger.debug('Process file as word 2007+ (docx)')
    return process_docx(filepath, field_filter_mode)


# === MAIN =================================================================


def process_maybe_encrypted(filepath, passwords=None, crypto_nesting=0,
                            **kwargs):
    """
    Process a file that might be encrypted.

    Calls :py:func:`process_file` and if that fails tries to decrypt and
    process the result. Based on recommendation in module doc string of
    :py:mod:`oletools.crypto`.

    :param str filepath: path to file on disc.
    :param passwords: list of passwords (str) to try for decryption or None
    :param int crypto_nesting: How many decryption layers were already used to
                               get the given file.
    :param kwargs: same as :py:func:`process_file`
    :returns: same as :py:func:`process_file`
    """
    # TODO: here filepath may also be a file in memory, it's not necessarily on disk
    result = u''
    try:
        result = process_file(filepath, **kwargs)
        if not crypto.is_encrypted(filepath):
            return result
    except Exception:
        logger.debug('Ignoring exception:', exc_info=True)
        if not crypto.is_encrypted(filepath):
            raise

    # we reach this point only if file is encrypted
    # check if this is an encrypted file in an encrypted file in an ...
    if crypto_nesting >= crypto.MAX_NESTING_DEPTH:
        raise crypto.MaxCryptoNestingReached(crypto_nesting, filepath)

    decrypted_file = None
    if passwords is None:
        passwords = crypto.DEFAULT_PASSWORDS
    else:
        passwords = list(passwords) + crypto.DEFAULT_PASSWORDS
    try:
        logger.debug('Trying to decrypt file')
        decrypted_file = crypto.decrypt(filepath, passwords)
        if not decrypted_file:
            logger.error('Decrypt failed, run with debug output to get details')
            raise crypto.WrongEncryptionPassword(filepath)
        logger.info('Analyze decrypted file')
        result = process_maybe_encrypted(decrypted_file, passwords,
                                         crypto_nesting+1, **kwargs)
    finally:     # clean up
        try:     # (maybe file was not yet created)
            os.unlink(decrypted_file)
        except Exception:
            logger.debug('Ignoring exception closing decrypted file:',
                         exc_info=True)
    return result


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
    log_helper.enable_logging(args.json, args.loglevel, stream=sys.stdout)

    if args.nounquote:
        global NO_QUOTES
        NO_QUOTES = True

    logger.print_str(BANNER)
    logger.print_str('Opening file: %s' % args.filepath)

    text = ''
    return_code = 1
    try:
        text = process_maybe_encrypted(
            args.filepath, args.password,
            field_filter_mode=args.field_filter_mode)
        return_code = 0
    except Exception as exc:
        logger.exception(str(exc))

    logger.print_str('DDE Links:')
    for link in text.splitlines():
        logger.print_str(text, type='dde-link')

    log_helper.end_logging()

    return return_code


if __name__ == '__main__':
    sys.exit(main())
