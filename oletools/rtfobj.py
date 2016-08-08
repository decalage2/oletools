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
# 2016-07-13 v0.50 PL: - new RtfParser and RtfObjParser classes
# 2016-07-18       SL: - added Python 3.5 support
# 2016-07-19       PL: - fixed Python 2.6-2.7 support
# 2016-07-30       PL: - new API with class RtfObject
#                      - backward-compatible API rtf_iter_objects (fixed issue #70)
# 2016-07-31       PL: - table output with tablestream
# 2016-08-01       PL: - detect executable filenames in OLE Package
# 2016-08-08       PL: - added option -s to save objects to files

__version__ = '0.50'

# ------------------------------------------------------------------------------
# TODO:
# - allow semicolon within hex, as found in  this sample:
#   http://contagiodump.blogspot.nl/2011/10/sep-28-cve-2010-3333-manuscript-with.html
# TODO: use OleObject and OleNativeStream in RtfObject instead of copying each attribute
# TODO: option -e <id> to extract an object, -e all for all objects
# TODO: option to choose which destinations to include (objdata by default)
# TODO: option to display SHA256 or MD5 hashes of objects in table


# === IMPORTS =================================================================

import re, os, sys, binascii, logging, optparse
import os.path

from thirdparty.xglob import xglob
from oleobj import OleObject, OleNativeStream
import oleobj

from thirdparty.tablestream import tablestream


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
HEX_DIGIT = b'[0-9A-Fa-f]'

# hex char = two alphanum digits: [0-9A-Fa-f]{2}
# HEX_CHAR = r'[0-9A-Fa-f]{2}'
# in fact MS Word allows whitespaces in between the hex digits!
# HEX_CHAR = r'[0-9A-Fa-f]\s*[0-9A-Fa-f]'
# Even worse, MS Word also allows ANY RTF-style tag {*} in between!!
# AND the tags can be nested...
#SINGLE_RTF_TAG = r'[{][^{}]*[}]'
# Actually RTF tags may contain braces escaped with backslash (\{ \}):
SINGLE_RTF_TAG = b'[{](?:\\\\.|[^{}\\\\])*[}]'

# Nested tags, two levels (because Python's re does not support nested matching):
# NESTED_RTF_TAG = r'[{](?:[^{}]|'+SINGLE_RTF_TAG+r')*[}]'
NESTED_RTF_TAG = b'[{](?:\\\\.|[^{}\\\\]|'+SINGLE_RTF_TAG+b')*[}]'

# AND it is also allowed to insert ANY control word or control symbol (ignored)
# According to Rich Text Format (RTF) Specification Version 1.9.1,
# section "Control Word":
# control word = \<ASCII Letter [a-zA-Z] Sequence max 32><Delimiter>
# delimiter = space, OR signed integer followed by any non-digit,
#             OR any character except letter and digit
# examples of valid control words:
# "\AnyThing " "\AnyThing123z" ""\AnyThing-456{" "\AnyThing{"
# control symbol = \<any char except letter or digit> (followed by anything)

ASCII_NAME = b'([a-zA-Z]{1,250})'

# using Python's re lookahead assumption:
# (?=...) Matches if ... matches next, but doesn't consume any of the string.
# This is called a lookahead assertion. For example, Isaac (?=Asimov) will
# match 'Isaac ' only if it's followed by 'Asimov'.

# TODO: Find the actual limit on the number of digits for Word
# SIGNED_INTEGER = r'(-?\d{1,250})'
SIGNED_INTEGER = b'(-?\\d+)'

CONTROL_WORD = b'(?:\\\\' + ASCII_NAME + b'(?:(?=[^a-zA-Z0-9-])|' + SIGNED_INTEGER + b'(?=[^0-9])))'

re_control_word = re.compile(CONTROL_WORD)

CONTROL_SYMBOL = b'(?:\\\\[^a-zA-Z0-9])'
re_control_symbol = re.compile(CONTROL_SYMBOL)

# Text that is not a control word/symbol or a group:
TEXT = b'[^{}\\\\]+'
re_text = re.compile(TEXT)

# ignored whitespaces and tags within a hex block:
IGNORED = b'(?:\\s|'+NESTED_RTF_TAG+b'|'+CONTROL_SYMBOL+b'|'+CONTROL_WORD+b')*'
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
PATTERN = b'\\b(?:' + HEX_DIGIT + IGNORED + b'){7,}' + HEX_DIGIT + b'\\b'

# at least 4 hex chars, followed by whitespace or CR/LF: (?:[0-9A-Fa-f]{2}){4,}\s*
# PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'
# improved pattern, allowing semicolons within hex:
#PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'

re_hexblock = re.compile(PATTERN)
re_embedded_tags = re.compile(IGNORED)
re_decimal = re.compile(b'\\d+')

re_delimiter = re.compile(b'[ \\t\\r\\n\\f\\v]')

DELIMITER = b'[ \\t\\r\\n\\f\\v]'
DELIMITERS_ZeroOrMore = b'[ \\t\\r\\n\\f\\v]*'
BACKSLASH_BIN = b'\\\\bin'
# According to my tests, Word accepts up to 250 digits (leading zeroes)
DECIMAL_GROUP = b'(\d{1,250})'

re_delims_bin_decimal = re.compile(DELIMITERS_ZeroOrMore + BACKSLASH_BIN
                                   + DECIMAL_GROUP + DELIMITER)
re_delim_hexblock = re.compile(DELIMITER + PATTERN)

# TODO: use a frozenset instead of a regex?
re_executable_extensions = re.compile(
    r"(?i)\.(EXE|COM|PIF|GADGET|MSI|MSP|MSC|VBS|VBE|VB|JSE|JS|WSF|WSC|WSH|WS|BAT|CMD|DLL|SCR|HTA|CPL|CLASS|JAR|PS1XML|PS1|PS2XML|PS2|PSC1|PSC2|SCF|LNK|INF|REG)\b")

# Destination Control Words, according to MS RTF Specifications v1.9.1:
DESTINATION_CONTROL_WORDS = frozenset((
    b"aftncn", b"aftnsep", b"aftnsepc", b"annotation", b"atnauthor", b"atndate", b"atnicn", b"atnid", b"atnparent", b"atnref",
    b"atntime", b"atrfend", b"atrfstart", b"author", b"background", b"bkmkend", b"bkmkstart", b"blipuid", b"buptim", b"category",
    b"colorschememapping", b"colortbl", b"comment", b"company", b"creatim", b"datafield", b"datastore", b"defchp", b"defpap",
    b"do", b"doccomm", b"docvar", b"dptxbxtext", b"ebcend", b"ebcstart", b"factoidname", b"falt", b"fchars", b"ffdeftext",
    b"ffentrymcr", b"ffexitmcr", b"ffformat", b"ffhelptext", b"ffl", b"ffname",b"ffstattext", b"field", b"file", b"filetbl",
    b"fldinst", b"fldrslt", b"fldtype", b"fname", b"fontemb", b"fontfile", b"fonttbl", b"footer", b"footerf", b"footerl",
    b"footerr", b"footnote", b"formfield", b"ftncn", b"ftnsep", b"ftnsepc", b"g", b"generator", b"gridtbl", b"header", b"headerf",
    b"headerl", b"headerr", b"hl", b"hlfr", b"hlinkbase", b"hlloc", b"hlsrc", b"hsv", b"htmltag", b"info", b"keycode", b"keywords",
    b"latentstyles", b"lchars", b"levelnumbers", b"leveltext", b"lfolevel", b"linkval", b"list", b"listlevel", b"listname",
    b"listoverride", b"listoverridetable", b"listpicture", b"liststylename", b"listtable", b"listtext", b"lsdlockedexcept",
    b"macc", b"maccPr", b"mailmerge", b"maln",b"malnScr", b"manager", b"margPr", b"mbar", b"mbarPr", b"mbaseJc", b"mbegChr",
    b"mborderBox", b"mborderBoxPr", b"mbox", b"mboxPr", b"mchr", b"mcount", b"mctrlPr", b"md", b"mdeg", b"mdegHide", b"mden",
    b"mdiff", b"mdPr", b"me", b"mendChr", b"meqArr", b"meqArrPr", b"mf", b"mfName", b"mfPr", b"mfunc", b"mfuncPr",b"mgroupChr",
    b"mgroupChrPr",b"mgrow", b"mhideBot", b"mhideLeft", b"mhideRight", b"mhideTop", b"mhtmltag", b"mlim", b"mlimloc", b"mlimlow",
    b"mlimlowPr", b"mlimupp", b"mlimuppPr", b"mm", b"mmaddfieldname", b"mmath", b"mmathPict", b"mmathPr",b"mmaxdist", b"mmc",
    b"mmcJc", b"mmconnectstr", b"mmconnectstrdata", b"mmcPr", b"mmcs", b"mmdatasource", b"mmheadersource", b"mmmailsubject",
    b"mmodso", b"mmodsofilter", b"mmodsofldmpdata", b"mmodsomappedname", b"mmodsoname", b"mmodsorecipdata", b"mmodsosort",
    b"mmodsosrc", b"mmodsotable", b"mmodsoudl", b"mmodsoudldata", b"mmodsouniquetag", b"mmPr", b"mmquery", b"mmr", b"mnary",
    b"mnaryPr", b"mnoBreak", b"mnum", b"mobjDist", b"moMath", b"moMathPara", b"moMathParaPr", b"mopEmu", b"mphant", b"mphantPr",
    b"mplcHide", b"mpos", b"mr", b"mrad", b"mradPr", b"mrPr", b"msepChr", b"mshow", b"mshp", b"msPre", b"msPrePr", b"msSub",
    b"msSubPr", b"msSubSup", b"msSubSupPr",  b"msSup", b"msSupPr", b"mstrikeBLTR", b"mstrikeH", b"mstrikeTLBR", b"mstrikeV",
    b"msub", b"msubHide", b"msup", b"msupHide", b"mtransp", b"mtype", b"mvertJc", b"mvfmf", b"mvfml", b"mvtof", b"mvtol",
    b"mzeroAsc", b"mzeroDesc", b"mzeroWid", b"nesttableprops", b"nexctfile", b"nonesttables", b"objalias", b"objclass",
    b"objdata", b"object", b"objname", b"objsect", b"objtime", b"oldcprops", b"oldpprops", b"oldsprops", b"oldtprops",
    b"oleclsid", b"operator", b"panose", b"password", b"passwordhash", b"pgp", b"pgptbl", b"picprop", b"pict", b"pn", b"pnseclvl",
    b"pntext", b"pntxta", b"pntxtb", b"printim", b"private", b"propname", b"protend", b"protstart", b"protusertbl", b"pxe",
    b"result", b"revtbl", b"revtim", b"rsidtbl", b"rtf", b"rxe", b"shp", b"shpgrp", b"shpinst", b"shppict", b"shprslt", b"shptxt",
    b"sn", b"sp", b"staticval", b"stylesheet", b"subject", b"sv", b"svb", b"tc", b"template", b"themedata", b"title", b"txe", b"ud",
    b"upr", b"userprops", b"wgrffmtfilter", b"windowcaption", b"writereservation", b"writereservhash", b"xe", b"xform",
    b"xmlattrname", b"xmlattrvalue", b"xmlclose", b"xmlname", b"xmlnstbl", b"xmlopen"
    ))


# some str methods on Python 2.x return characters,
# while the equivalent bytes methods return integers on Python 3.x:
if sys.version_info[0] <= 2:
    # Python 2.x - Characters (str)
    BACKSLASH = '\\'
    BRACE_OPEN = '{'
    BRACE_CLOSE = '}'
else:
    # Python 3.x - Integers
    BACKSLASH = ord('\\')
    BRACE_OPEN = ord('{')
    BRACE_CLOSE = ord('}')


#=== CLASSES =================================================================

class Destination(object):
    """
    Stores the data associated with a destination control word
    """
    def __init__(self, cword=None):
        self.cword = cword
        self.data = b''
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
            if self.data[self.index] == BRACE_OPEN:
                self._open_group()
                self.index += 1
                continue
            if self.data[self.index] == BRACE_CLOSE:
                self._close_group()
                self.index += 1
                continue
            if self.data[self.index] == BACKSLASH:
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
                    if cword == b'bin':
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
        #log.debug('{ Open Group at index %Xh - level=%d' % (self.index, self.group_level))
        # call user method AFTER increasing the level:
        self.open_group()

    def open_group(self):
        #log.debug('open group at index %Xh' % self.index)
        pass

    def _close_group(self):
        #log.debug('} Close Group at index %Xh - level=%d' % (self.index, self.group_level))
        # call user method BEFORE decreasing the level:
        self.close_group()
        # if the destination level is the same as the group level, close the destination:
        if self.group_level == self.current_destination.group_level:
            # log.debug('Current Destination %r level = %d => Close Destination' % (
            #     self.current_destination.cword, self.current_destination.group_level))
            self._close_destination()
        else:
            # log.debug('Current Destination %r level = %d => Continue with same Destination' % (
            #     self.current_destination.cword, self.current_destination.group_level))
            pass
        self.group_level -= 1
        # log.debug('Decreased group level to %d' % self.group_level)

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
        # log.debug("Open Destination %r start=%Xh - level=%d" % (cword, new_dest.start, new_dest.group_level))
        # call the corresponding user method for additional processing:
        self.open_destination(self.current_destination)

    def open_destination(self, destination):
        pass

    def _close_destination(self):
        # log.debug("Close Destination %r end=%Xh - level=%d" % (self.current_destination.cword,
        #     self.index, self.current_destination.group_level))
        self.current_destination.end = self.index
        # call the corresponding user method for additional processing:
        self.close_destination(self.current_destination)
        if len(self.destinations)>0:
            # remove the current destination from the stack, and go back to the previous one:
            self.destinations.pop()
        if len(self.destinations) > 0:
            self.current_destination = self.destinations[-1]
        else:
            # log.debug('All destinations are closed, keeping the document destination open')
            pass

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
        # log.debug('%Xh Reached End of File')
        # close any group/destination that is still open:
        while self.group_level > 0:
            # log.debug('Group Level = %d, closing group' % self.group_level)
            self._close_group()
        self.end_of_file()

    def end_of_file(self):
        pass


class RtfObject(object):
    """
    An object or a file (OLE Package) embedded into an RTF document
    """
    def __init__(self):
        """
        RtfObject constructor
        """
        # start and end index in the RTF file:
        self.start = None
        self.end = None
        # raw object data encoded in hexadecimal, as found in the RTF file:
        self.hexdata = None
        # raw object data in binary form, decoded from hexadecimal
        self.rawdata = None
        # OLE object data (extracted from rawdata)
        self.is_ole = False
        self.oledata = None
        self.format_id = None
        self.class_name = None
        self.oledata_size = None
        # OLE Package data (extracted from oledata)
        self.is_package = False
        self.olepkgdata = None
        self.filename = None
        self.src_path = None
        self.temp_path = None



class RtfObjParser(RtfParser):
    """
    Specialized RTF parser to extract OLE objects
    """

    def __init__(self, data):
        super(RtfObjParser, self).__init__(data)
        # list of RtfObjects found
        self.objects = []

    def open_destination(self, destination):
        if destination.cword == b'objdata':
            log.debug('*** Start object data at index %Xh' % destination.start)

    def close_destination(self, destination):
        if destination.cword == b'objdata':
            log.debug('*** Close object data at index %Xh' % self.index)
            rtfobj = RtfObject()
            self.objects.append(rtfobj)
            rtfobj.start = destination.start
            rtfobj.end = destination.end
            # Filter out all whitespaces first (just ignored):
            hexdata1 = destination.data.translate(None, b' \t\r\n\f\v')
            # Then filter out any other non-hex character:
            hexdata = re.sub(b'[^a-hA-H0-9]', b'', hexdata1)
            if len(hexdata) < len(hexdata1):
                # this is only for debugging:
                nonhex = re.sub(b'[a-hA-H0-9]', b'', hexdata1)
                log.debug('Found non-hex chars in hexdata: %r' % nonhex)
            # MS Word accepts an extra hex digit, so we need to trim it if present:
            if len(hexdata) & 1:
                log.debug('Odd length, trimmed last byte.')
                hexdata = hexdata[:-1]
            rtfobj.hexdata = hexdata
            object_data = binascii.unhexlify(hexdata)
            rtfobj.rawdata = object_data
            # TODO: check if all hex data is extracted properly

            obj = OleObject()
            try:
                obj.parse(object_data)
                rtfobj.format_id = obj.format_id
                rtfobj.class_name = obj.class_name
                rtfobj.oledata_size = obj.data_size
                rtfobj.oledata = obj.data
                rtfobj.is_ole = True
                if obj.class_name.lower() == 'package':
                    opkg = OleNativeStream(bindata=obj.data)
                    rtfobj.filename = opkg.filename
                    rtfobj.src_path = opkg.src_path
                    rtfobj.temp_path = opkg.temp_path
                    rtfobj.olepkgdata = opkg.data
                    rtfobj.is_package = True
            except:
                pass
                log.debug('*** Not an OLE 1.0 Object')

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

def rtf_iter_objects(filename, min_size=32):
    """
    [DEPRECATED] Backward-compatible API, for applications using the old rtfobj:
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    data = open(filename, 'rb').read()
    rtfp = RtfObjParser(data)
    rtfp.parse()
    for obj in rtfp.objects:
        # orig_len = obj.end - obj.start
        yield obj.start, obj.rawdata





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


def process_file(container, filename, data, output_dir=None, save_object=False):
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
    print('='*79)
    print('File: %r - size: %d bytes' % (filename, len(data)))
    tstream = tablestream.TableStream(
        column_width=(3, 10, 31, 31),
        header_row=('id', 'index', 'OLE Object', 'OLE Package'),
        style=tablestream.TableStyleSlim
    )
    rtfp = RtfObjParser(data)
    rtfp.parse()
    for rtfobj in rtfp.objects:
        pkg_color = None
        if rtfobj.is_ole:
            ole_column = 'format_id: %d\n' % rtfobj.format_id
            ole_column += 'class name: %r\n' % rtfobj.class_name
            ole_column += 'data size: %d' % rtfobj.oledata_size
            if rtfobj.is_package:
                pkg_column = 'Filename: %r\n' % rtfobj.filename
                pkg_column += 'Source path: %r\n' % rtfobj.src_path
                pkg_column += 'Temp path = %r' % rtfobj.temp_path
                pkg_color = 'yellow'
                # check if the file extension is executable:
                _, ext = os.path.splitext(rtfobj.filename)
                log.debug('File extension: %r' % ext)
                if re_executable_extensions.match(ext):
                    pkg_color = 'red'
                    pkg_column += '\nEXECUTABLE FILE'
            else:
                pkg_column = 'Not an OLE Package'
        else:
            pkg_column = ''
            ole_column = 'Not a well-formed OLE object'
        tstream.write_row((
            rtfp.objects.index(rtfobj),
            # filename,
            '%08Xh' % rtfobj.start,
            ole_column,
            pkg_column
            ), colors=(None, None, None, pkg_color)
        )
        tstream.write_sep()
    if save_object:
        if save_object == 'all':
            objects = rtfp.objects
        else:
            try:
                i = int(save_object)
                objects = [ rtfp.objects[i] ]
            except:
                log.error('The -s option must be followed by an object index or all, such as "-s 2" or "-s all"')
                return
        for rtfobj in objects:
            i = objects.index(rtfobj)
            if rtfobj.is_package:
                print('Saving file from OLE Package in object #%d:' % i)
                print('  Filename = %r' % rtfobj.filename)
                print('  Source path = %r' % rtfobj.src_path)
                print('  Temp path = %r' % rtfobj.temp_path)
                if rtfobj.filename:
                    fname = '%s_%s' % (fname_prefix,
                                       sanitize_filename(rtfobj.filename))
                else:
                    fname = '%s_object_%08X.noname' % (fname_prefix, rtfobj.start)
                print('  saving to file %s' % fname)
                open(fname, 'wb').write(rtfobj.olepkgdata)
            elif rtfobj.is_ole:
                print('Saving file embedded in OLE object #%d:' % i)
                print('  format_id  = %d' % rtfobj.format_id)
                print('  class name = %r' % rtfobj.class_name)
                print('  data size  = %d' % rtfobj.oledata_size)
                # set a file extension according to the class name:
                class_name = rtfobj.class_name.lower()
                if class_name.startswith(b'word'):
                    ext = 'doc'
                elif class_name.startswith(b'package'):
                    ext = 'package'
                else:
                    ext = 'bin'
                fname = '%s_object_%08X.%s' % (fname_prefix, rtfobj.start, ext)
                print('  saving to file %s' % fname)
                open(fname, 'wb').write(rtfobj.oledata)
            else:
                print('Saving raw data in object #%d:' % i)
                fname = '%s_object_%08X.raw' % (fname_prefix, rtfobj.start)
                print('  saving object to file %s' % fname)
                open(fname, 'wb').write(rtfobj.rawdata)


#=== MAIN =================================================================

def main():
    # print banner with version
    print ('rtfobj %s - http://decalage.info/python/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')

    DEFAULT_LOG_LEVEL = "warning" # Default log level
    LOG_LEVELS = {
        'debug':    logging.DEBUG,
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
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
        help='if the file is a zip archive, open first file from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
        help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")
    parser.add_option("-s", "--save", dest='save_object', type='str', default=None,
        help='Save the object corresponding to the provided number to a file, for example "-s 2". Use "-s all" to save all objects at once.')
    # parser.add_option("-o", "--outfile", dest='outfile', type='str', default=None,
    #     help='Filename to be used when saving an object to a file.')
    parser.add_option("-d", type="str", dest="output_dir",
        help='use specified directory to save output files.', default=None)
    # parser.add_option("--pkg", action="store_true", dest="save_pkg",
    #     help='Save OLE Package binary data of extracted objects (file embedded into an OLE Package).')
    # parser.add_option("--ole", action="store_true", dest="save_ole",
    #     help='Save OLE binary data of extracted objects (object data without the OLE container).')
    # parser.add_option("--raw", action="store_true", dest="save_raw",
    #     help='Save raw binary data of extracted objects (decoded from hex, including the OLE container).')
    # parser.add_option("--hex", action="store_true", dest="save_hex",
    #     help='Save raw hexadecimal data of extracted objects (including the OLE container).')


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
        process_file(container, filename, data, output_dir=options.output_dir,
                     save_object=options.save_object)


if __name__ == '__main__':
    main()

# This code was developed while listening to The Mary Onettes "Lost"

