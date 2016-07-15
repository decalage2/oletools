#!/usr/bin/env python
from __future__ import print_function

"""
rtfobj.py

rtfobj is a Python module to extract embedded objects from RTF files, such as
OLE ojects. It can be used as a Python library or a command-line tool.

Usage: rtfobj.py <file.rtf>

rtfobj project website: http://www.decalage.info/python/rtfobj

rtfobj is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# rtfobj is copyright (c) 2012-2016, Philippe Lagadec (http://www.decalage.info)
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


#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-11-09 v0.01 PL: - first version
# 2013-04-02 v0.02 PL: - fixed bug in main
# 2015-12-09 v0.03 PL: - configurable logging, CLI options
#                      - extract OLE 1.0 objects
#                      - extract files from OLE Package objects
# 2016-04-01 v0.04 PL: - fixed logging output to use stdout instead of stderr
# 2016-04-07 v0.45 PL: - improved parsing to handle some malware tricks
# 2016-05-06 v0.47 TJ: - added option -d to set the output directory
#                        (contribution by Thomas Jarosch)
#                  TJ: - sanitize filenames to avoid special characters
# 2016-05-29       PL: - improved parsing, fixed issue #42
# 2016-07-13 v0.48 PL: - new RtfParser and RtfObjParser classes

__version__ = '0.48'

#------------------------------------------------------------------------------
# TODO:
# - allow semicolon within hex, as found in  this sample:
#   http://contagiodump.blogspot.nl/2011/10/sep-28-cve-2010-3333-manuscript-with.html


#=== IMPORTS =================================================================

import re, os, sys, string, binascii, logging, optparse

from thirdparty.xglob import xglob
from oleobj import OleObject, OleNativeStream
import oleobj


# === LOGGING =================================================================

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
log = get_logger('rtfobj')


#=== CONSTANTS=================================================================

# REGEX pattern to extract embedded OLE objects in hexadecimal format:

# alphanum digit: [0-9A-Fa-f]
HEX_DIGIT = r'[0-9A-Fa-f]'

# hex char = two alphanum digits: [0-9A-Fa-f]{2}
# HEX_CHAR = r'[0-9A-Fa-f]{2}'
# in fact MS Word allows whitespaces in between the hex digits!
# HEX_CHAR = r'[0-9A-Fa-f]\s*[0-9A-Fa-f]'
# Even worse, MS Word also allows ANY RTF-style tag {*} in between!!
# AND the tags can be nested...
#SINGLE_RTF_TAG = r'[{][^{}]*[}]'
# Actually RTF tags may contain braces escaped with backslash (\{ \}):
SINGLE_RTF_TAG = r'[{](?:\\.|[^{}\\])*[}]'

# Nested tags, two levels (because Python's re does not support nested matching):
# NESTED_RTF_TAG = r'[{](?:[^{}]|'+SINGLE_RTF_TAG+r')*[}]'
NESTED_RTF_TAG = r'[{](?:\\.|[^{}\\]|'+SINGLE_RTF_TAG+r')*[}]'

# AND it is also allowed to insert ANY control word or control symbol (ignored)
# According to Rich Text Format (RTF) Specification Version 1.9.1,
# section "Control Word":
# control word = \<ASCII Letter [a-zA-Z] Sequence max 32><Delimiter>
# delimiter = space, OR signed integer followed by any non-digit,
#             OR any character except letter and digit
# examples of valid control words:
# "\AnyThing " "\AnyThing123z" ""\AnyThing-456{" "\AnyThing{"
# control symbol = \<any char except letter or digit> (followed by anything)

ASCII_NAME = r'([a-zA-Z]{1,250})'

# using Python's re lookahead assumption:
# (?=...) Matches if ... matches next, but doesn't consume any of the string.
# This is called a lookahead assertion. For example, Isaac (?=Asimov) will
# match 'Isaac ' only if it's followed by 'Asimov'.

# TODO: Find the actual limit on the number of digits for Word
# SIGNED_INTEGER = r'(-?\d{1,250})'
SIGNED_INTEGER = r'(-?\d+)'

CONTROL_WORD = r'(?:\\' + ASCII_NAME + r'(?:(?=[^a-zA-Z0-9-])|' + SIGNED_INTEGER + r'(?=[^0-9])))'
re_control_word = re.compile(CONTROL_WORD)

CONTROL_SYMBOL = r'(?:\\[^a-zA-Z0-9])'
re_control_symbol = re.compile(CONTROL_SYMBOL)

# Text that is not a control word/symbol or a group:
TEXT = r'[^{}\\]+'
re_text = re.compile(TEXT)

# ignored whitespaces and tags within a hex block:
IGNORED = r'(?:\s|'+NESTED_RTF_TAG+'|'+CONTROL_SYMBOL+'|'+CONTROL_WORD+r')*'
#IGNORED = r'\s*'

# HEX_CHAR = HEX_DIGIT + IGNORED + HEX_DIGIT

# several hex chars, at least 4: (?:[0-9A-Fa-f]{2}){4,}
# + word boundaries
# HEX_CHARS_4orMORE = r'\b(?:' + HEX_CHAR + r'){4,}\b'
# at least 1 hex char:
# HEX_CHARS_1orMORE = r'(?:' + HEX_CHAR + r')+'
# at least 1 hex char, followed by whitespace or CR/LF:
# HEX_CHARS_1orMORE_WHITESPACES = r'(?:' + HEX_CHAR + r')+\s+'
# + word boundaries around hex block
# HEX_CHARS_1orMORE_WHITESPACES = r'\b(?:' + HEX_CHAR + r')+\b\s*'
# at least one block of hex and whitespace chars, followed by closing curly bracket:
# HEX_BLOCK_CURLY_BRACKET = r'(?:' + HEX_CHARS_1orMORE_WHITESPACES + r')+\}'
# PATTERN = r'(?:' + HEX_CHARS_1orMORE_WHITESPACES + r')*' + HEX_CHARS_1orMORE

#TODO PATTERN = r'\b(?:' + HEX_CHAR + IGNORED + r'){4,}\b'
# PATTERN = r'\b(?:' + HEX_CHAR + IGNORED + r'){4,}' #+ HEX_CHAR + r'\b'
PATTERN = r'\b(?:' + HEX_DIGIT + IGNORED + r'){7,}' + HEX_DIGIT + r'\b'

# at least 4 hex chars, followed by whitespace or CR/LF: (?:[0-9A-Fa-f]{2}){4,}\s*
# PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'
# improved pattern, allowing semicolons within hex:
#PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'

# a dummy translation table for str.translate, which does not change anythying:
TRANSTABLE_NOCHANGE = string.maketrans('', '')

re_hexblock = re.compile(PATTERN)
re_embedded_tags = re.compile(IGNORED)
re_decimal = re.compile(r'\d+')

re_delimiter = re.compile(r'[ \t\r\n\f\v]')

DELIMITER = r'[ \t\r\n\f\v]'
DELIMITERS_ZeroOrMore = r'[ \t\r\n\f\v]*'
BACKSLASH_BIN = r'\\bin'
# According to my tests, Word accepts up to 250 digits (leading zeroes)
DECIMAL_GROUP = r'(\d{1,250})'

re_delims_bin_decimal = re.compile(DELIMITERS_ZeroOrMore + BACKSLASH_BIN
                                   + DECIMAL_GROUP + DELIMITER)
re_delim_hexblock = re.compile(DELIMITER + PATTERN)

# Destination Control Words, according to MS RTF Specifications v1.9.1:
DESTINATION_CONTROL_WORDS = frozenset((
    "aftncn", "aftnsep", "aftnsepc", "annotation", "atnauthor", "atndate", "atnicn", "atnid", "atnparent", "atnref",
    "atntime", "atrfend", "atrfstart", "author", "background", "bkmkend", "bkmkstart", "blipuid", "buptim", "category",
    "colorschememapping", "colortbl", "comment", "company", "creatim", "datafield", "datastore", "defchp", "defpap",
    "do", "doccomm", "docvar", "dptxbxtext", "ebcend", "ebcstart", "factoidname", "falt", "fchars", "ffdeftext",
    "ffentrymcr", "ffexitmcr", "ffformat", "ffhelptext", "ffl", "ffname", "ffstattext", "field", "file", "filetbl",
    "fldinst", "fldrslt", "fldtype", "fname", "fontemb", "fontfile", "fonttbl", "footer", "footerf", "footerl",
    "footerr", "footnote", "formfield", "ftncn", "ftnsep", "ftnsepc", "g", "generator", "gridtbl", "header", "headerf",
    "headerl", "headerr", "hl", "hlfr", "hlinkbase", "hlloc", "hlsrc", "hsv", "htmltag", "info", "keycode", "keywords",
    "latentstyles", "lchars", "levelnumbers", "leveltext", "lfolevel", "linkval", "list", "listlevel", "listname",
    "listoverride", "listoverridetable", "listpicture", "liststylename", "listtable", "listtext", "lsdlockedexcept",
    "macc", "maccPr", "mailmerge", "maln", "malnScr", "manager", "margPr", "mbar", "mbarPr", "mbaseJc", "mbegChr",
    "mborderBox", "mborderBoxPr", "mbox", "mboxPr", "mchr", "mcount", "mctrlPr", "md", "mdeg", "mdegHide", "mden",
    "mdiff", "mdPr", "me", "mendChr", "meqArr", "meqArrPr", "mf", "mfName", "mfPr", "mfunc", "mfuncPr", "mgroupChr",
    "mgroupChrPr", "mgrow", "mhideBot", "mhideLeft", "mhideRight", "mhideTop", "mhtmltag", "mlim", "mlimloc", "mlimlow",
    "mlimlowPr", "mlimupp", "mlimuppPr", "mm", "mmaddfieldname", "mmath", "mmathPict", "mmathPr", "mmaxdist", "mmc",
    "mmcJc", "mmconnectstr", "mmconnectstrdata", "mmcPr", "mmcs", "mmdatasource", "mmheadersource", "mmmailsubject",
    "mmodso", "mmodsofilter", "mmodsofldmpdata", "mmodsomappedname", "mmodsoname", "mmodsorecipdata", "mmodsosort",
    "mmodsosrc", "mmodsotable", "mmodsoudl", "mmodsoudldata", "mmodsouniquetag", "mmPr", "mmquery", "mmr", "mnary",
    "mnaryPr", "mnoBreak", "mnum", "mobjDist", "moMath", "moMathPara", "moMathParaPr", "mopEmu", "mphant", "mphantPr",
    "mplcHide", "mpos", "mr", "mrad", "mradPr", "mrPr", "msepChr", "mshow", "mshp", "msPre", "msPrePr", "msSub",
    "msSubPr", "msSubSup", "msSubSupPr", "msSup", "msSupPr", "mstrikeBLTR", "mstrikeH", "mstrikeTLBR", "mstrikeV",
    "msub", "msubHide", "msup", "msupHide", "mtransp", "mtype", "mvertJc", "mvfmf", "mvfml", "mvtof", "mvtol",
    "mzeroAsc", "mzeroDesc", "mzeroWid", "nesttableprops", "nextfile", "nonesttables", "objalias", "objclass",
    "objdata", "object", "objname", "objsect", "objtime", "oldcprops", "oldpprops", "oldsprops", "oldtprops",
    "oleclsid", "operator", "panose", "password", "passwordhash", "pgp", "pgptbl", "picprop", "pict", "pn", "pnseclvl",
    "pntext", "pntxta", "pntxtb", "printim", "private", "propname", "protend", "protstart", "protusertbl", "pxe",
    "result", "revtbl", "revtim", "rsidtbl", "rtf", "rxe", "shp", "shpgrp", "shpinst", "shppict", "shprslt", "shptxt",
    "sn", "sp", "staticval", "stylesheet", "subject", "sv", "svb", "tc", "template", "themedata", "title", "txe", "ud",
    "upr", "userprops", "wgrffmtfilter", "windowcaption", "writereservation", "writereservhash", "xe", "xform",
    "xmlattrname", "xmlattrvalue", "xmlclose", "xmlname", "xmlnstbl", "xmlopen"
    ))



#=== CLASSES =================================================================

class Destination(object):
    """
    Stores the data associated with a destination control word
    """
    def __init__(self, cword=None):
        self.cword = cword
        self.data = ''
        self.start = None
        self.end = None
        self.group_level = 0


# class Group(object):
#     """
#     Stores the data associated with a group between braces {...}
#     """
#     def __init__(self, cword=None):
#         self.start = None
#         self.end = None
#         self.level = None



class RtfParser(object):
    """
    Very simple generic RTF parser
    """

    def __init__(self, data):
        self.data = data
        self.index = 0
        self.size = len(data)
        self.group_level = 0
        # default destination for the document text:
        document_destination = Destination()
        self.destinations = [document_destination]
        self.current_destination = document_destination

    def parse(self):
        self.index = 0
        while self.index < self.size:
            if self.data[self.index] == '{':
                self._open_group()
                self.index += 1
                continue
            if self.data[self.index] == '}':
                self._close_group()
                self.index += 1
                continue
            if self.data[self.index] == '\\':
                m = re_control_word.match(self.data, self.index)
                if m:
                    cword = m.group(1)
                    param = None
                    if len(m.groups()) > 1:
                        param = m.group(2)
                    # log.debug('control word %r at index %Xh - cword=%r param=%r' % (m.group(), self.index, cword, param))
                    self._control_word(m, cword, param)
                    self.index += len(m.group())
                    # if it's \bin, call _bin after updating index
                    if cword == 'bin':
                        self._bin(m, param)
                    continue
                m = re_control_symbol.match(self.data, self.index)
                if m:
                    self.control_symbol(m)
                    self.index += len(m.group())
                    continue
            m = re_text.match(self.data, self.index)
            if m:
                self._text(m)
                self.index += len(m.group())
                continue
            raise RuntimeError('Should not have reached this point - index=%Xh' % self.index)
        self.end_of_file()


    def _open_group(self):
        self.group_level += 1
        log.debug('{ Open Group at index %Xh - level=%d' % (self.index, self.group_level))
        # call user method AFTER increasing the level:
        self.open_group()

    def open_group(self):
        #log.debug('open group at index %Xh' % self.index)
        pass

    def _close_group(self):
        log.debug('} Close Group at index %Xh - level=%d' % (self.index, self.group_level))
        # call user method BEFORE decreasing the level:
        self.close_group()
        # if the destination level is the same as the group level, close the destination:
        if self.group_level == self.current_destination.group_level:
            log.debug('Current Destination %r level = %d => Close Destination' % (
                self.current_destination.cword, self.current_destination.group_level))
            self._close_destination()
        else:
            log.debug('Current Destination %r level = %d => Continue with same Destination' % (
                self.current_destination.cword, self.current_destination.group_level))
        self.group_level -= 1
        log.debug('Decreased group level to %d' % self.group_level)

    def close_group(self):
        #log.debug('close group at index %Xh' % self.index)
        pass

    def _open_destination(self, matchobject, cword):
        # if the current destination is at the same group level, close it first:
        if self.current_destination.group_level == self.group_level:
            self._close_destination()
        new_dest = Destination(cword)
        new_dest.group_level = self.group_level
        self.destinations.append(new_dest)
        self.current_destination = new_dest
        # start of the destination is right after the control word:
        new_dest.start = self.index + len(matchobject.group())
        log.debug("Open Destination %r start=%Xh - level=%d" % (cword, new_dest.start, new_dest.group_level))
        # call the corresponding user method for additional processing:
        self.open_destination(self.current_destination)

    def open_destination(self, destination):
        pass

    def _close_destination(self):
        log.debug("Close Destination %r end=%Xh - level=%d" % (self.current_destination.cword,
            self.index, self.current_destination.group_level))
        self.current_destination.end = self.index
        # call the corresponding user method for additional processing:
        self.close_destination(self.current_destination)
        if len(self.destinations)>0:
            # remove the current destination from the stack, and go back to the previous one:
            self.destinations.pop()
        if len(self.destinations) > 0:
            self.current_destination = self.destinations[-1]
        else:
            log.debug('All destinations are closed, keeping the document destination open')

    def close_destination(self, destination):
        pass

    def _control_word(self, matchobject, cword, param):
        #log.debug('control word %r at index %Xh' % (matchobject.group(), self.index))
        if cword in DESTINATION_CONTROL_WORDS:
            # log.debug('%r is a destination control word: starting a new destination' % cword)
            self._open_destination(matchobject, cword)
        # call the corresponding user method for additional processing:
        self.control_word(matchobject, cword, param)

    def control_word(self, matchobject, cword, param):
        pass

    def control_symbol(self, matchobject):
        #log.debug('control symbol %r at index %Xh' % (matchobject.group(), self.index))
        pass

    def _text(self, matchobject):
        text = matchobject.group()
        self.current_destination.data += text
        self.text(matchobject, text)

    def text(self, matchobject, text):
        #log.debug('text %r at index %Xh' % (matchobject.group(), self.index))
        pass

    def _bin(self, matchobject, param):
        binlen = int(param)
        log.debug('\\bin: reading %d bytes of binary data' % binlen)
        # TODO: handle optional space?
        # TODO: handle negative length, and length greater than data
        bindata = self.data[self.index:self.index + binlen]
        self.index += binlen
        self.bin(bindata)

    def bin(self, bindata):
        pass

    def _end_of_file(self):
        log.debug('%Xh Reached End of File')
        # close any group/destination that is still open:
        while self.group_level > 0:
            log.debug('Group Level = %d, closing group' % self.group_level)
            self._close_group()
        self.end_of_file()

    def end_of_file(self):
        pass


class RtfObjParser(RtfParser):
    """
    Specialized RTF parser to extract OLE objects
    """

    def __init__(self, data, fname_prefix='rtf'):
        super(RtfObjParser, self).__init__(data)
        self.fname_prefix = fname_prefix

    def open_destination(self, destination):
        if destination.cword == 'objdata':
            log.debug('*** Start object data at index %Xh' % destination.start)

    def close_destination(self, destination):
        if destination.cword == 'objdata':
            log.debug('*** Close object data at index %Xh' % self.index)
            # Filter out all whitespaces first (just ignored):
            hexdata1 = destination.data.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
            # Then filter out any other non-hex character:
            hexdata = re.sub(r'[^a-hA-H0-9]', '', hexdata1)
            if len(hexdata) < len(hexdata1):
                # this is only for debugging:
                nonhex = re.sub(r'[a-hA-H0-9]', '', hexdata1)
                log.debug('Found non-hex chars in hexdata: %r' % nonhex)
            # MS Word accepts an extra hex digit, so we need to trim it if present:
            if len(hexdata) & 1:
                log.debug('Odd length, trimmed last byte.')
                hexdata = hexdata[:-1]
            object_data = binascii.unhexlify(hexdata)
            print('found object size %d at index %08X - end %08X' % (len(object_data),
                                                                     destination.start, self.index))
            fname = '%s_object_%08X.raw' % (self.fname_prefix, destination.start)
            print('saving object to file %s' % fname)
            open(fname, 'wb').write(object_data)
            # TODO: check if all hex data is extracted properly

            obj = OleObject()
            try:
                obj.parse(object_data)
                print('extract file embedded in OLE object:')
                print('format_id  = %d' % obj.format_id)
                print('class name = %r' % obj.class_name)
                print('data size  = %d' % obj.data_size)
                # set a file extension according to the class name:
                class_name = obj.class_name.lower()
                if class_name.startswith('word'):
                    ext = 'doc'
                elif class_name.startswith('package'):
                    ext = 'package'
                else:
                    ext = 'bin'

                fname = '%s_object_%08X.%s' % (self.fname_prefix, destination.start, ext)
                print('saving to file %s' % fname)
                open(fname, 'wb').write(obj.data)
                if obj.class_name.lower() == 'package':
                    print('Parsing OLE Package')
                    opkg = OleNativeStream(bindata=obj.data)
                    print('Filename = %r' % opkg.filename)
                    print('Source path = %r' % opkg.src_path)
                    print('Temp path = %r' % opkg.temp_path)
                    if opkg.filename:
                        fname = '%s_%s' % (self.fname_prefix,
                                           sanitize_filename(opkg.filename))
                    else:
                        fname = '%s_object_%08X.noname' % (self.fname_prefix, destination.start)
                    print('saving to file %s' % fname)
                    open(fname, 'wb').write(opkg.data)
            except:
                pass
                log.exception('*** Not an OLE 1.0 Object')

    def bin(self, bindata):
        if self.current_destination.cword == 'objdata':
            # TODO: keep track of this, because it is unusual and indicates potential obfuscation
            # trick: hexlify binary data, add it to hex data
            self.current_destination.data += binascii.hexlify(bindata)

    def control_word(self, matchobject, cword, param):
        # TODO: extract useful cwords such as objclass
        # TODO: keep track of cwords inside objdata, because it is unusual and indicates potential obfuscation
        # TODO: same with control symbols, and opening bracket
        pass


#=== FUNCTIONS ===============================================================

def rtf_iter_objects_old (filename, min_size=32):
    """
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    data = open(filename, 'rb').read()
    for m in re.finditer(PATTERN, data):
        found = m.group(0)
        orig_len = len(found)
        # remove all whitespace and line feeds:
        #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
        found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v}')
        found = binascii.unhexlify(found)
        #print repr(found)
        if len(found)>min_size:
            yield m.start(), orig_len, found

# TODO: backward-compatible API?


def search_hex_block(data, pos=0, min_size=32, first=True):
    if first:
        # Search 1st occurence of a hex block:
        match = re_hexblock.search(data, pos=pos)
    else:
        # Match next occurences of a hex block, from the current position only:
        match = re_hexblock.match(data, pos=pos)



def rtf_iter_objects (data, min_size=32):
    """
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    # Search 1st occurence of a hex block:
    match = re_hexblock.search(data)
    if match is None:
        log.debug('No hex block found.')
        # no hex block found
        return
    while match is not None:
        found = match.group(0)
        # start index
        start = match.start()
        # current position
        current = match.end()
        log.debug('Found hex block starting at %08X, end %08X, size=%d' % (start, current, len(found)))
        if len(found) < min_size:
            log.debug('Too small - size<%d, ignored.' % min_size)
            match = re_hexblock.search(data, pos=current)
            continue
        #log.debug('Match: %s' % found)
        # remove all whitespace and line feeds:
        #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
        found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
        # TODO: make it a function
        # Also remove embedded RTF tags:
        found = re_embedded_tags.sub('', found)
        # object data extracted from the RTF file
        # MS Word accepts an extra hex digit, so we need to trim it if present:
        if len(found) & 1:
            log.debug('Odd length, trimmed last byte.')
            found = found[:-1]
        #log.debug('Cleaned match: %s' % found)
        objdata = binascii.unhexlify(found)
        # Detect the "\bin" control word, which is sometimes used for obfuscation:
        bin_match = re_delims_bin_decimal.match(data, pos=current)
        while bin_match is not None:
            log.debug('Found \\bin block starting at %08X : %r'
                          % (bin_match.start(), bin_match.group(0)))
            # extract the decimal integer following '\bin'
            bin_len = int(bin_match.group(1))
            log.debug('\\bin block length = %d' % bin_len)
            if current+bin_len > len(data):
                log.error('\\bin block length is larger than the remaining data')
                # move the current index, ignore the \bin block
                current += len(bin_match.group(0))
                break
            # read that number of bytes:
            objdata += data[current:current+bin_len]
            # TODO: handle exception
            current += len(bin_match.group(0)) + bin_len
            # TODO: check if current is out of range
            # TODO: is Word limiting the \bin length to a number of digits?
            log.debug('Current position = %08X' % current)
            match = re_delim_hexblock.match(data, pos=current)
            if match is not None:
                log.debug('Found next hex block starting at %08X, end %08X'
                    % (match.start(), match.end()))
                found = match.group(0)
                log.debug('Match: %s' % found)
                # remove all whitespace and line feeds:
                #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
                found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
                # Also remove embedded RTF tags:
                found = re_embedded_tags.sub(found, '')
                objdata += binascii.unhexlify(found)
                current = match.end()
            bin_match = re_delims_bin_decimal.match(data, pos=current)

        # print repr(found)
        if len(objdata)>min_size:
            yield start, current-start, objdata
        # Search next occurence of a hex block:
        match = re_hexblock.search(data, pos=current)



def sanitize_filename(filename, replacement='_', max_length=200):
    """compute basename of filename. Replaces all non-whitelisted characters.
       The returned filename is always a basename of the file."""
    basepath = os.path.basename(filename).strip()
    sane_fname = re.sub(r'[^\w\.\- ]', replacement, basepath)

    while ".." in sane_fname:
        sane_fname = sane_fname.replace('..', '.')

    while "  " in sane_fname:
        sane_fname = sane_fname.replace('  ', ' ')

    if not len(filename):
        sane_fname = 'NONAME'

    # limit filename length
    if max_length:
        sane_fname = sane_fname[:max_length]

    return sane_fname


def process_file(container, filename, data, output_dir=None):
    if output_dir:
        if not os.path.isdir(output_dir):
            log.info('creating output directory %s' % output_dir)
            os.mkdir(output_dir)

        fname_prefix = os.path.join(output_dir,
                                    sanitize_filename(filename))
    else:
        base_dir = os.path.dirname(filename)
        sane_fname = sanitize_filename(filename)
        fname_prefix = os.path.join(base_dir, sane_fname)

    # TODO: option to extract objects to files (false by default)
    if data is None:
        data = open(filename, 'rb').read()
    rtfp = RtfObjParser(data, fname_prefix)
    rtfp.parse()

    # print '-'*79
    # print 'File: %r - %d bytes' % (filename, len(data))
    # for index, orig_len, objdata in rtf_iter_objects(data):
    #     print 'found object size %d at index %08X - end %08X' % (len(objdata), index, index+orig_len)
    #     fname = '%s_object_%08X.raw' % (fname_prefix, index)
    #     print 'saving object to file %s' % fname
    #     open(fname, 'wb').write(objdata)
    #     # TODO: check if all hex data is extracted properly
    #
    #     obj = OleObject()
    #     try:
    #         obj.parse(objdata)
    #         print 'extract file embedded in OLE object:'
    #         print 'format_id  = %d' % obj.format_id
    #         print 'class name = %r' % obj.class_name
    #         print 'data size  = %d' % obj.data_size
    #         # set a file extension according to the class name:
    #         class_name = obj.class_name.lower()
    #         if class_name.startswith('word'):
    #             ext = 'doc'
    #         elif class_name.startswith('package'):
    #             ext = 'package'
    #         else:
    #             ext = 'bin'
    #
    #         fname = '%s_object_%08X.%s' % (fname_prefix, index, ext)
    #         print 'saving to file %s' % fname
    #         open(fname, 'wb').write(obj.data)
    #         if obj.class_name.lower() == 'package':
    #             print 'Parsing OLE Package'
    #             opkg = OleNativeStream(bindata=obj.data)
    #             print 'Filename = %r' % opkg.filename
    #             print 'Source path = %r' % opkg.src_path
    #             print 'Temp path = %r' % opkg.temp_path
    #             if opkg.filename:
    #                 fname = '%s_%s' % (fname_prefix,
    #                                    sanitize_filename(opkg.filename))
    #             else:
    #                 fname = '%s_object_%08X.noname' % (fname_prefix, index)
    #             print 'saving to file %s' % fname
    #             open(fname, 'wb').write(opkg.data)
    #     except:
    #         pass
    #         log.exception('*** Not an OLE 1.0 Object')



#=== MAIN =================================================================

if __name__ == '__main__':
    # print banner with version
    print ('rtfobj %s - http://decalage.info/python/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')

    DEFAULT_LOG_LEVEL = "warning" # Default log level
    LOG_LEVELS = {'debug':    logging.DEBUG,
              'info':     logging.INFO,
              'warning':  logging.WARNING,
              'error':    logging.ERROR,
              'critical': logging.CRITICAL
             }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    # parser.add_option('-o', '--outfile', dest='outfile',
    #     help='output file')
    # parser.add_option('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_option("-r", action="store_true", dest="recursive",
        help='find files recursively in subdirectories.')
    parser.add_option("-d", type="str", dest="output_dir",
        help='use specified directory to output files.', default=None)
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
        help='if the file is a zip archive, open first file from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
        help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print (__doc__)
        parser.print_help()
        sys.exit()

    # Setup logging to the console:
    # here we use stdout instead of stderr by default, so that the output
    # can be redirected properly.
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)
    oleobj.log.setLevel(logging.NOTSET)


    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
        zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        process_file(container, filename, data, options.output_dir)


# This code was developed while listening to The Mary Onettes "Lost"

