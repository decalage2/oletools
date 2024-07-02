#!/usr/bin/env python
"""
olevba.py

olevba is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to extract VBA Macro code in clear text, deobfuscate
and analyze malicious macros.
XLM/Excel 4 Macros are also supported in Excel and SLK files.

Supported formats:
    - Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
    - Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
    - PowerPoint 97-2003 (.ppt), PowerPoint 2007+ (.pptm, .ppsm)
    - Word/PowerPoint 2007+ XML (aka Flat OPC)
    - Word 2003 XML (.xml)
    - Word/Excel Single File Web Page / MHTML (.mht)
    - Publisher (.pub)
    - SYLK/SLK files (.slk)
    - Text file containing VBA or VBScript source code
    - Password-protected Zip archive containing any of the above
    - raises an error if run with files encrypted using MS Crypto API RC4

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

olevba is part of the python-oletools package:
http://www.decalage.info/python/oletools

olevba is based on source code from officeparser by John William Davison
https://github.com/unixfreak0037/officeparser
"""

# === LICENSE ==================================================================

# olevba is copyright (c) 2014-2024 Philippe Lagadec (http://www.decalage.info)
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


# olevba contains modified source code from the officeparser project, published
# under the following MIT License (MIT):
#
# officeparser is copyright (c) 2014 John William Davison
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2014-08-05 v0.01 PL: - first version based on officeparser code
# 2014-08-14 v0.02 PL: - fixed bugs in code, added license from officeparser
# 2014-08-15       PL: - fixed incorrect value check in projecthelpfilepath Record
# 2014-08-15 v0.03 PL: - refactored extract_macros to support OpenXML formats
#                        and to find the VBA project root anywhere in the file
# 2014-11-29 v0.04 PL: - use olefile instead of OleFileIO_PL
# 2014-12-05 v0.05 PL: - refactored most functions into a class, new API
#                      - added detect_vba_macros
# 2014-12-10 v0.06 PL: - hide first lines with VB attributes
#                      - detect auto-executable macros
#                      - ignore empty macros
# 2014-12-14 v0.07 PL: - detect_autoexec() is now case-insensitive
# 2014-12-15 v0.08 PL: - improved display for empty macros
#                      - added pattern extraction
# 2014-12-25 v0.09 PL: - added suspicious keywords detection
# 2014-12-27 v0.10 PL: - added OptionParser, main and process_file
#                      - uses xglob to scan several files with wildcards
#                      - option -r to recurse subdirectories
#                      - option -z to scan files in password-protected zips
# 2015-01-02 v0.11 PL: - improved filter_vba to detect colons
# 2015-01-03 v0.12 PL: - fixed detect_patterns to detect all patterns
#                      - process_file: improved display, shows container file
#                      - improved list of executable file extensions
# 2015-01-04 v0.13 PL: - added several suspicious keywords, improved display
# 2015-01-08 v0.14 PL: - added hex strings detection and decoding
#                      - fixed issue #2, decoding VBA stream names using
#                        specified codepage and unicode stream names
# 2015-01-11 v0.15 PL: - added new triage mode, options -t and -d
# 2015-01-16 v0.16 PL: - fix for issue #3 (exception when module name="text")
#                      - added several suspicious keywords
#                      - added option -i to analyze VBA source code directly
# 2015-01-17 v0.17 PL: - removed .com from the list of executable extensions
#                      - added scan_vba to run all detection algorithms
#                      - decoded hex strings are now also scanned + reversed
# 2015-01-23 v0.18 PL: - fixed issue #3, case-insensitive search in code_modules
# 2015-01-24 v0.19 PL: - improved the detection of IOCs obfuscated with hex
#                        strings and StrReverse
# 2015-01-26 v0.20 PL: - added option --hex to show all hex strings decoded
# 2015-01-29 v0.21 PL: - added Dridex obfuscation decoding
#                      - improved display, shows obfuscation name
# 2015-02-01 v0.22 PL: - fixed issue #4: regex for URL, e-mail and exe filename
#                      - added Base64 obfuscation decoding (contribution from
#                        @JamesHabben)
# 2015-02-03 v0.23 PL: - triage now uses VBA_Scanner results, shows Base64 and
#                        Dridex strings
#                      - exception handling in detect_base64_strings
# 2015-02-07 v0.24 PL: - renamed option --hex to --decode, fixed display
#                      - display exceptions with stack trace
#                      - added several suspicious keywords
#                      - improved Base64 detection and decoding
#                      - fixed triage mode not to scan attrib lines
# 2015-03-04 v0.25 PL: - added support for Word 2003 XML
# 2015-03-22 v0.26 PL: - added suspicious keywords for sandboxing and
#                        virtualisation detection
# 2015-05-06 v0.27 PL: - added support for MHTML files with VBA macros
#                        (issue #10 reported by Greg from SpamStopsHere)
# 2015-05-24 v0.28 PL: - improved support for MHTML files with modified header
#                        (issue #11 reported by Thomas Chopitea)
# 2015-05-26 v0.29 PL: - improved MSO files parsing, taking into account
#                        various data offsets (issue #12)
#                      - improved detection of MSO files, avoiding incorrect
#                        parsing errors (issue #7)
# 2015-05-29 v0.30 PL: - added suspicious keywords suggested by @ozhermit,
#                        Davy Douhine (issue #9), issue #13
# 2015-06-16 v0.31 PL: - added generic VBA expression deobfuscation (chr,asc,etc)
# 2015-06-19       PL: - added options -a, -c, --each, --attr
# 2015-06-21 v0.32 PL: - always display decoded strings which are printable
#                      - fix VBA_Scanner.scan to return raw strings, not repr()
# 2015-07-09 v0.40 PL: - removed usage of sys.stderr which causes issues
# 2015-07-12       PL: - added Hex function decoding to VBA Parser
# 2015-07-13       PL: - added Base64 function decoding to VBA Parser
# 2015-09-06       PL: - improved VBA_Parser, refactored the main functions
# 2015-09-13       PL: - moved main functions to a class VBA_Parser_CLI
#                      - fixed issue when analysis was done twice
# 2015-09-15       PL: - remove duplicate IOCs from results
# 2015-09-16       PL: - join long VBA lines ending with underscore before scan
#                      - disabled unused option --each
# 2015-09-22 v0.41 PL: - added new option --reveal
#                      - added suspicious strings for PowerShell.exe options
# 2015-10-09 v0.42 PL: - VBA_Parser: split each format into a separate method
# 2015-10-10       PL: - added support for text files with VBA source code
# 2015-11-17       PL: - fixed bug with --decode option
# 2015-12-16       PL: - fixed bug in main (no options input anymore)
#                      - improved logging, added -l option
# 2016-01-31       PL: - fixed issue #31 in VBA_Parser.open_mht
#                      - fixed issue #32 by monkeypatching email.feedparser
# 2016-02-07       PL: - KeyboardInterrupt is now raised properly
# 2016-02-20 v0.43 PL: - fixed issue #34 in the VBA parser and vba_chr
# 2016-02-29       PL: - added Workbook_Activate to suspicious keywords
# 2016-03-08 v0.44 PL: - added VBA Form strings extraction and analysis
# 2016-03-04 v0.45 CH: - added JSON output (by Christian Herdtweck)
# 2016-03-16       CH: - added option --no-deobfuscate (temporary)
# 2016-04-19 v0.46 PL: - new option --deobf instead of --no-deobfuscate
#                      - updated suspicious keywords
# 2016-05-04 v0.47 PL: - look for VBA code in any stream including orphans
# 2016-04-28       CH: - return an exit code depending on the results
#                      - improved error and exception handling
#                      - improved JSON output
# 2016-05-12       CH: - added support for PowerPoint 97-2003 files
# 2016-06-06       CH: - improved handling of unicode VBA module names
# 2016-06-07       CH: - added option --relaxed, stricter parsing by default
# 2016-06-12 v0.50 PL: - fixed small bugs in VBA parsing code
# 2016-07-01       PL: - fixed issue #58 with format() to support Python 2.6
# 2016-07-29       CH: - fixed several bugs including #73 (Mac Roman encoding)
# 2016-08-31       PL: - added autoexec keyword InkPicture_Painted
#                      - detect_autoexec now returns the exact keyword found
# 2016-09-05       PL: - added autoexec keywords for MS Publisher (.pub)
# 2016-09-06       PL: - fixed issue #20, is_zipfile on Python 2.6
# 2016-09-12       PL: - enabled packrat to improve pyparsing performance
# 2016-10-25       PL: - fixed raise and print statements for Python 3
# 2016-11-03 v0.51 PL: - added EnumDateFormats and EnumSystemLanguageGroupsW
# 2017-02-07       PL: - temporary fix for issue #132
#                      - added keywords for Mac-specific macros (issue #130)
# 2017-03-08       PL: - fixed absolute imports
# 2017-03-16       PL: - fixed issues #148 and #149 for option --reveal
# 2017-05-19       PL: - added enable_logging to fix issue #154
# 2017-05-31     c1fe: - PR #135 fixing issue #132 for some Mac files
# 2017-06-08       PL: - fixed issue #122 Chr() with negative numbers
# 2017-06-15       PL: - deobfuscation line by line to handle large files
# 2017-07-11 v0.52 PL: - raise exception instead of sys.exit (issue #180)
# 2017-11-08       VB: - PR #124 adding user form parsing (Vincent Brillault)
# 2017-11-17       PL: - fixed a few issues with form parsing
# 2017-11-20       PL: - fixed issue #219, do not close the file too early
# 2017-11-24       PL: - added keywords to detect self-modifying macros and
#                        attempts to disable macro security (issue #221)
# 2018-03-19       PL: - removed pyparsing from the thirdparty subfolder
# 2018-04-15 v0.53 PL: - added support for Word/PowerPoint 2007+ XML (FlatOPC)
#                        (issue #283)
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2018-10-08       PL: - replace backspace before printing to console (issue #358)
# 2018-10-25       CH: - detect encryption and raise error if detected
# 2018-12-03       PL: - uses tablestream (+colors) instead of prettytable
# 2018-12-06       PL: - colorize the suspicious keywords found in VBA code
# 2019-01-01       PL: - removed support for Python 2.6
# 2019-03-18       PL: - added XLM/XLF macros detection for Excel OLE files
# 2019-03-25       CH: - added decryption of password-protected files
# 2019-04-09       PL: - decompress_stream accepts bytes (issue #422)
# 2019-05-23 v0.55 PL: - added option --pcode to call pcodedmp and display P-code
# 2019-06-05       PL: - added VBA stomping detection
# 2019-09-24       PL: - included DridexUrlDecode into olevba (issue #485)
# 2019-12-03       PL: - added support for SLK files and XLM macros in SLK
# 2020-01-31 v0.56 KS: - added option --no-xlm, improved MHT detection
# 2020-03-22       PL: - uses plugin_biff to display DCONN objects and their URL
# 2020-06-11       PL: - fixed issue #575 when decompressing raw chunks in VBA
# 2020-09-03       MX: - fixed issue #602 monkeypatch in email package
# 2020-09-16       PL: - enabled relaxed mode by default (issues #477, #593)
#                      - fixed detect_vba_macros to always return VBA code as
#                        unicode on Python 3 (issues  #455, #477, #587, #593)
# 2020-09-28       PL: - added VBA_Parser.get_vba_code_all_modules (partial fix
#                        for issue #619)
# 2021-04-14       PL: - added detection of Workbook_BeforeClose (issue #518)
# 2021-11-09       KJ: - added PROJECTCOMPATVERSION Record on dir Stream

__version__ = '0.60.2'

#------------------------------------------------------------------------------
# TODO:
# + add xor bruteforcing like bbharvest
# + options -a and -c should imply -d

# TODO later:
# + performance improvement: instead of searching each keyword separately,
#   first split vba code into a list of words (per line), then check each
#   word against a dict. (or put vba words into a set/dict?)
# + for regex, maybe combine them into a single re with named groups?
# + add Yara support, include sample rules? plugins like balbuzard?
# + add balbuzard support
# + output to file (replace print by file.write, sys.stdout by default)
# + look for VBA in embedded documents (e.g. Excel in Word)
# + support SRP streams (see Lenny's article + links and sample)
# - python 3.x support
# - check VBA macros in Visio, Access, Project, etc
# - extract_macros: convert to a class, split long function into smaller methods
# - extract_macros: read bytes from stream file objects instead of strings
# - extract_macros: use combined struct.unpack instead of many calls
# - all except clauses should target specific exceptions

# ------------------------------------------------------------------------------
# REFERENCES:
# - [MS-OVBA]: Microsoft Office VBA File Format Structure
#   http://msdn.microsoft.com/en-us/library/office/cc313094%28v=office.12%29.aspx
# - officeparser: https://github.com/unixfreak0037/officeparser


# --- IMPORTS ------------------------------------------------------------------

import traceback
import sys
import os
import struct
from io import BytesIO, StringIO
import math
import zipfile
import re
import argparse
import binascii
import base64
import zlib
import email  # for MHTML parsing
import email.feedparser
import string  # for printable
import json   # for json output mode (argument --json)

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    try:
        # Python 2.5+: batteries included
        import xml.etree.cElementTree as ET
    except ImportError:
        try:
            # Python <2.5: standalone ElementTree install
            import elementtree.cElementTree as ET
        except ImportError:
            raise ImportError("lxml or ElementTree are not installed, " \
                               + "see http://codespeak.net/lxml " \
                               + "or http://effbot.org/zone/element-index.htm")

import colorclass

# On Windows, colorclass needs to be enabled:
if os.name == 'nt':
    colorclass.Windows.enable(auto_colors=True)

from pyparsing import \
        CaselessKeyword, CaselessLiteral, Combine, Forward, Literal, \
        Optional, QuotedString,Regex, Suppress, Word, WordStart, \
        alphanums, alphas, hexnums,nums, opAssoc, srange, \
        infixNotation, ParserElement

# attempt to import XLMMacroDeobfuscator (optional)
try:
    from XLMMacroDeobfuscator import deobfuscator as xlmdeobfuscator
    XLMDEOBFUSCATOR = True
except ImportError:
    XLMDEOBFUSCATOR = False

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

import olefile
from oletools.thirdparty.tablestream import tablestream
from oletools.thirdparty.xglob import xglob, PathNotFoundException
from oletools.thirdparty.oledump.plugin_biff import cBIFF
from oletools import ppt_parser
from oletools import oleform
from oletools import rtfobj
from oletools import crypto
from oletools.common.io_encoding import ensure_stdout_handles_unicode
from oletools.common import codepages
from oletools import ftguess
from oletools.common.log_helper import log_helper

# === PYTHON 2+3 SUPPORT ======================================================

if sys.version_info[0] <= 2:
    # Python 2.x
    PYTHON2 = True
    # to use ord on bytes/bytearray items the same way in Python 2+3
    # on Python 2, just use the normal ord() because items are bytes
    byte_ord = ord
    #: Default string encoding for the olevba API
    DEFAULT_API_ENCODING = 'utf8'  # on Python 2: UTF-8 (bytes)
else:
    # Python 3.x+
    PYTHON2 = False

    # to use ord on bytes/bytearray items the same way in Python 2+3
    # on Python 3, items are int, so just return the item
    def byte_ord(x):
        return x
    # xrange is now called range:
    xrange = range
    # unichr does not exist anymore, only chr:
    unichr = chr
    # json2ascii also needs "unicode":
    unicode = str
    from functools import reduce
    #: Default string encoding for the olevba API
    DEFAULT_API_ENCODING = None  # on Python 3: None (unicode)
    # Python 3.0 - 3.4 support:
    # From https://gist.github.com/ynkdir/867347/c5e188a4886bc2dd71876c7e069a7b00b6c16c61
    if sys.version_info < (3, 5):
        import codecs
        _backslashreplace_errors = codecs.lookup_error("backslashreplace")

        def backslashreplace_errors(exc):
            if isinstance(exc, UnicodeDecodeError):
                u = "".join("\\x{0:02x}".format(c) for c in exc.object[exc.start:exc.end])
                return u, exc.end
            return _backslashreplace_errors(exc)

        codecs.register_error("backslashreplace", backslashreplace_errors)


def unicode2str(unicode_string):
    """
    convert a unicode string to a native str:
        - on Python 3, it returns the same string
        - on Python 2, the string is encoded with UTF-8 to a bytes str
    :param unicode_string: unicode string to be converted
    :return: the string converted to str
    :rtype: str
    """
    if PYTHON2:
        return unicode_string.encode('utf8', errors='replace')
    else:
        return unicode_string


def bytes2str(bytes_string, encoding='utf8'):
    """
    convert a bytes string to a native str:
        - on Python 2, it returns the same string (bytes=str)
        - on Python 3, the string is decoded using the provided encoding
          (UTF-8 by default) to a unicode str
    :param bytes_string: bytes string to be converted
    :param encoding: codec to be used for decoding
    :return: the string converted to str
    :rtype: str
    """
    if PYTHON2:
        return bytes_string
    else:
        return bytes_string.decode(encoding, errors='replace')


# === LOGGING =================================================================


# a global logger object used for debugging:
log = log_helper.get_or_create_silent_logger('olevba')


def enable_logging():
    """
    Enable logging for this module (disabled by default).

    For use by third-party libraries that import `olevba` as module.

    This will set the module-specific logger level to `NOTSET`, which
    means the main application controls the actual logging level.

    This also enables logging for the modules used by us, but not the global
    common logging mechanism (:py:mod:`oletools.common.log_helper.log_helper`).
    Use :py:func:`oletools.common.log_helper.log_helper.enable_logging` for
    that.
    """
    log.setLevel(log_helper.NOTSET)
    ppt_parser.enable_logging()
    crypto.enable_logging()
    # TODO: do not have enable_logging yet: oleform, rtfobj


#=== EXCEPTIONS ==============================================================

class OlevbaBaseException(Exception):
    """ Base class for exceptions produced here for simpler except clauses """
    def __init__(self, msg, filename=None, orig_exc=None, **kwargs):
        if orig_exc:
            super(OlevbaBaseException, self).__init__(msg +
                                                      ' ({0})'.format(orig_exc),
                                                      **kwargs)
        else:
            super(OlevbaBaseException, self).__init__(msg, **kwargs)
        self.msg = msg
        self.filename = filename
        self.orig_exc = orig_exc


class FileOpenError(OlevbaBaseException):
    """ raised by VBA_Parser constructor if all open_... attempts failed

    probably means the file type is not supported
    """

    def __init__(self, filename, orig_exc=None):
        super(FileOpenError, self).__init__(
            'Failed to open file %s' % filename, filename, orig_exc)


class ProcessingError(OlevbaBaseException):
    """ raised by VBA_Parser.process_file* functions """

    def __init__(self, filename, orig_exc):
        super(ProcessingError, self).__init__(
            'Error processing file %s' % filename, filename, orig_exc)


class MsoExtractionError(RuntimeError, OlevbaBaseException):
    """ raised by mso_file_extract if parsing MSO/ActiveMIME data failed """

    def __init__(self, msg):
        MsoExtractionError.__init__(self, msg)
        OlevbaBaseException.__init__(self, msg)


class SubstreamOpenError(FileOpenError):
    """ special kind of FileOpenError: file is a substream of original file """

    def __init__(self, filename, subfilename, orig_exc=None):
        super(SubstreamOpenError, self).__init__(
            str(filename) + '/' + str(subfilename), orig_exc)
        self.filename = filename   # overwrite setting in OlevbaBaseException
        self.subfilename = subfilename


class UnexpectedDataError(OlevbaBaseException):
    """ raised when parsing is strict (=not relaxed) and data is unexpected """

    def __init__(self, stream_path, variable, expected, value):
        if isinstance(expected, int):
            es = '{0:04X}'.format(expected)
        elif isinstance(expected, tuple):
            es = ','.join('{0:04X}'.format(e) for e in expected)
            es =  '({0})'.format(es)
        else:
            raise ValueError('Unknown type encountered: {0}'.format(type(expected)))
        super(UnexpectedDataError, self).__init__(
            'Unexpected value in {0} for variable {1}: '
            'expected {2} but found {3:04X}!'
            .format(stream_path, variable, es, value))
        self.stream_path = stream_path
        self.variable = variable
        self.expected = expected
        self.value = value

#--- CONSTANTS ----------------------------------------------------------------

# return codes
RETURN_OK             = 0
RETURN_WARNINGS       = 1  # (reserved, not used yet)
RETURN_WRONG_ARGS     = 2  # (fixed, built into argparse)
RETURN_FILE_NOT_FOUND = 3
RETURN_XGLOB_ERR      = 4
RETURN_OPEN_ERROR     = 5
RETURN_PARSE_ERROR    = 6
RETURN_SEVERAL_ERRS   = 7
RETURN_UNEXPECTED     = 8
RETURN_ENCRYPTED      = 9

# MAC codepages (from http://stackoverflow.com/questions/1592925/decoding-mac-os-text-in-python)
MAC_CODEPAGES = {
    10000: 'mac-roman',
    10001: 'shiftjis',  # not found: 'mac-shift-jis',
    10003: 'ascii',     # nothing appropriate found: 'mac-hangul',
    10008: 'gb2321',    # not found: 'mac-gb2312',
    10002: 'big5',      # not found: 'mac-big5',
    10005: 'hebrew',    # not found: 'mac-hebrew',
    10004: 'mac-arabic',
    10006: 'mac-greek',
    10081: 'mac-turkish',
    10021: 'thai',      # not found: mac-thai',
    10029: 'maccentraleurope',  # not found: 'mac-east europe',
    10007: 'ascii',     # nothing appropriate found: 'mac-russian',
}

# URL and message to report issues:
URL_OLEVBA_ISSUES = 'https://github.com/decalage2/oletools/issues'
MSG_OLEVBA_ISSUES = 'Please report this issue on %s' % URL_OLEVBA_ISSUES

# Container types:
TYPE_OLE = 'OLE'
TYPE_OpenXML = 'OpenXML'
TYPE_FlatOPC_XML = 'FlatOPC_XML'
TYPE_Word2003_XML = 'Word2003_XML'
TYPE_MHTML = 'MHTML'
TYPE_TEXT = 'Text'
TYPE_PPT = 'PPT'
TYPE_SLK = 'SLK'

# short tag to display file types in triage mode:
TYPE2TAG = {
    TYPE_OLE: 'OLE:',
    TYPE_OpenXML: 'OpX:',
    TYPE_FlatOPC_XML: 'FlX:',
    TYPE_Word2003_XML: 'XML:',
    TYPE_MHTML: 'MHT:',
    TYPE_TEXT: 'TXT:',
    TYPE_PPT: 'PPT:',
    TYPE_SLK: 'SLK:',
}


# MSO files ActiveMime header magic
MSO_ACTIVEMIME_HEADER = b'ActiveMime'

MODULE_EXTENSION = "bas"
CLASS_EXTENSION = "cls"
FORM_EXTENSION = "frm"

# Namespaces and tags for Word2003 XML parsing:
NS_W = '{http://schemas.microsoft.com/office/word/2003/wordml}'
# the tag <w:binData w:name="editdata.mso"> contains the VBA macro code:
TAG_BINDATA = NS_W + 'binData'
ATTR_NAME = NS_W + 'name'

# Namespaces and tags for Word/PowerPoint 2007+ XML parsing:
# root: <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
NS_XMLPACKAGE = '{http://schemas.microsoft.com/office/2006/xmlPackage}'
TAG_PACKAGE = NS_XMLPACKAGE + 'package'
# the tag <pkg:part> includes <pkg:binaryData> that contains the VBA macro code in Base64:
# <pkg:part pkg:name="/word/vbaProject.bin" pkg:contentType="application/vnd.ms-office.vbaProject"><pkg:binaryData>
TAG_PKGPART = NS_XMLPACKAGE + 'part'
ATTR_PKG_NAME = NS_XMLPACKAGE + 'name'
ATTR_PKG_CONTENTTYPE = NS_XMLPACKAGE + 'contentType'
CTYPE_VBAPROJECT = "application/vnd.ms-office.vbaProject"
TAG_PKGBINDATA = NS_XMLPACKAGE + 'binaryData'

# Keywords to detect auto-executable macros
# Simple strings, without regex characters:
AUTOEXEC_KEYWORDS = {
    # MS Word:
    'Runs when the Word document is opened':
        ('AutoExec', 'AutoOpen', 'DocumentOpen'),
    'Runs when the Word document is closed':
        ('AutoExit', 'AutoClose', 'Document_Close', 'DocumentBeforeClose'),
    'Runs when the Word document is modified':
        ('DocumentChange',),
    'Runs when a new Word document is created':
        ('AutoNew', 'Document_New', 'NewDocument'),

    # MS Word and Publisher:
    'Runs when the Word or Publisher document is opened':
        ('Document_Open',),
    'Runs when the Publisher document is closed':
        ('Document_BeforeClose',),

    # MS Excel:
    'Runs when the Excel Workbook is opened':
        ('Auto_Open', 'Workbook_Open', 'Workbook_Activate', 'Auto_Ope'),
        # TODO: "Auto_Ope" is temporarily here because of a bug in plugin_biff, which misses the last byte in "Auto_Open"...
    'Runs when the Excel Workbook is closed':
        ('Auto_Close', 'Workbook_Close', 'Workbook_BeforeClose'),
        #Worksheet_Calculate to Autoexec: see http://www.certego.net/en/news/advanced-vba-macros/
    'May run when an Excel WorkSheet is opened':
        ('Worksheet_Calculate',),
}

# Keywords to detect auto-executable macros
# Regular expressions:
AUTOEXEC_KEYWORDS_REGEX = {
    # any MS Office application:
    'Runs when the file is opened (using InkPicture ActiveX object)':
        # ref:https://twitter.com/joe4security/status/770691099988025345
        (r'\w+_Painted', r'\w+_Painting'),
    'Runs when the file is opened and ActiveX objects trigger events':
        (r'\w+_GotFocus', r'\w+_LostFocus', r'\w+_MouseHover', r'\w+_Click',
         r'\w+_Change', r'\w+_Resize', r'\w+_BeforeNavigate2', r'\w+_BeforeScriptExecute',
         r'\w+_DocumentComplete', r'\w+_DownloadBegin', r'\w+_DownloadComplete',
         r'\w+_FileDownload', r'\w+_NavigateComplete2', r'\w+_NavigateError',
         r'\w+_ProgressChange', r'\w+_PropertyChange', r'\w+_SetSecureLockIcon',
         r'\w+_StatusTextChange', r'\w+_TitleChange', r'\w+_MouseMove', r'\w+_MouseEnter',
         r'\w+_MouseLeave', r'\w+_Layout', r'\w+_OnConnecting', r'\w+_FollowHyperlink', r'\w+_ContentControlOnEnter'),
}

# Suspicious Keywords that may be used by malware
# See VBA language reference: http://msdn.microsoft.com/en-us/library/office/jj692818%28v=office.15%29.aspx
SUSPICIOUS_KEYWORDS = {
    #TODO: use regex to support variable whitespaces
    #http://www.certego.net/en/news/advanced-vba-macros/
    'May read system environment variables':
        ('Environ','Win32_Environment','Environment','ExpandEnvironmentStrings','HKCU\\Environment',
        'HKEY_CURRENT_USER\\Environment'),
    'May open a file':
        ('Open',),
    'May write to a file (if combined with Open)':
    #TODO: regex to find Open+Write on same line
        ('Write', 'Put', 'Output', 'Print #'),
    'May read or write a binary file (if combined with Open)':
    #TODO: regex to find Open+Binary on same line
        ('Binary',),
    'May copy a file':
        ('FileCopy', 'CopyFile','CopyHere','CopyFolder'),
    #FileCopy: http://msdn.microsoft.com/en-us/library/office/gg264390%28v=office.15%29.aspx
    #CopyFile: http://msdn.microsoft.com/en-us/library/office/gg264089%28v=office.15%29.aspx
    #CopyHere, MoveHere, MoveHere and MoveFolder exploitation: see http://www.certego.net/en/news/advanced-vba-macros/
    'May move a file':
        ('MoveHere', 'MoveFile', 'MoveFolder'),
    'May delete a file':
        ('Kill',),
    'May create a text file':
        ('CreateTextFile', 'ADODB.Stream', 'WriteText', 'SaveToFile'),
    #CreateTextFile: http://msdn.microsoft.com/en-us/library/office/gg264617%28v=office.15%29.aspx
    #ADODB.Stream sample: http://pastebin.com/Z4TMyuq6
    #ShellExecute: https://twitter.com/StanHacked/status/1075088449768693762
    #InvokeVerb, InvokeVerbEx, DoIt and ControlPanelItem: see http://www.certego.net/en/news/advanced-vba-macros/

    'May run an executable file or a system command':
        ('Shell', 'vbNormal', 'vbNormalFocus', 'vbHide', 'vbMinimizedFocus', 'vbMaximizedFocus', 'vbNormalNoFocus',
         'vbMinimizedNoFocus', 'WScript.Shell', 'Run', 'ShellExecute', 'ShellExecuteA', 'shell32','InvokeVerb','InvokeVerbEx',
         'DoIt'),
    'May run a dll':
        ('ControlPanelItem',),
    # Win32_Process.Create https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/create-method-in-class-win32-process
    'May execute file or a system command through WMI':
        ('Create',),
    # WMI https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/create-method-in-class-win32-process
    # MacScript: see https://msdn.microsoft.com/en-us/library/office/gg264812.aspx
    # AppleScript: see https://docs.microsoft.com/en-us/office/vba/office-mac/applescripttask
    'May run an executable file or a system command on a Mac':
        ('MacScript','AppleScript'),
    #Shell: http://msdn.microsoft.com/en-us/library/office/gg278437%28v=office.15%29.aspx
    #WScript.Shell+Run sample: http://pastebin.com/Z4TMyuq6
    'May run PowerShell commands':
    #sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
    #also: https://bitbucket.org/decalage/oletools/issues/14/olevba-library-update-ioc
    # ref: https://blog.netspi.com/15-ways-to-bypass-the-powershell-execution-policy/
    # TODO: add support for keywords starting with a non-alpha character, such as "-noexit"
    # TODO: '-command', '-EncodedCommand', '-scriptblock'
        ('PowerShell', 'noexit', 'ExecutionPolicy', 'noprofile', 'command', 'EncodedCommand',
         'invoke-command', 'scriptblock', 'Invoke-Expression', 'AuthorizationManager'),
    'May run an executable file or a system command using PowerShell':
        ('Start-Process',),
    'May call a DLL using Excel 4 Macros (XLM/XLF)':
        ('CALL',),
    'May hide the application':
        ('Application.Visible', 'ShowWindow', 'SW_HIDE'),
    'May create a directory':
        ('MkDir',),
    'May save the current workbook':
        ('ActiveWorkbook.SaveAs',),
    'May change which directory contains files to open at startup':
    #TODO: confirm the actual effect
        ('Application.AltStartupPath',),
    'May create an OLE object':
        ('CreateObject',),
    #bypass CreateObject http://www.certego.net/en/news/advanced-vba-macros/
    'May get an OLE object with a running instance':
        ('GetObject',),
    'May create an OLE object using PowerShell':
        ('New-Object',),
    'May run an application (if combined with CreateObject)':
        ('Shell.Application',),
    'May run an Excel 4 Macro (aka XLM/XLF) from VBA':
        ('ExecuteExcel4Macro',),
    'May enumerate application windows (if combined with Shell.Application object)':
        ('Windows', 'FindWindow'),
    'May run code from a DLL':
    #TODO: regex to find declare+lib on same line - see mraptor
        ('Lib',),
    'May run code from a library on a Mac':
    #TODO: regex to find declare+lib on same line - see mraptor
        ('libc.dylib', 'dylib'),
    'May inject code into another process':
        ('CreateThread', 'CreateUserThread', 'VirtualAlloc', # (issue #9) suggested by Davy Douhine - used by MSF payload
        'VirtualAllocEx', 'RtlMoveMemory', 'WriteProcessMemory',
        'SetContextThread', 'QueueApcThread', 'WriteVirtualMemory', 'VirtualProtect',
        ),
    'May run a shellcode in memory':
        ('SetTimer',  # Vidar sample: https://app.any.run/tasks/897f28e7-3162-4b65-b268-2655543199d6/
         ),
    'May download files from the Internet':
    #TODO: regex to find urlmon+URLDownloadToFileA on same line
        ('URLDownloadToFileA', 'Msxml2.XMLHTTP', 'Microsoft.XMLHTTP',
         'MSXML2.ServerXMLHTTP', # suggested in issue #13
         'User-Agent', # sample from @ozhermit: http://pastebin.com/MPc3iV6z
        ),
    'May download files from the Internet using PowerShell':
    #sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
        ('Net.WebClient', 'DownloadFile', 'DownloadString'),
    'May control another application by simulating user keystrokes':
        ('SendKeys', 'AppActivate'),
    #SendKeys: http://msdn.microsoft.com/en-us/library/office/gg278655%28v=office.15%29.aspx
    'May attempt to obfuscate malicious function calls':
        ('CallByName',),
    #CallByName: http://msdn.microsoft.com/en-us/library/office/gg278760%28v=office.15%29.aspx
    'May attempt to obfuscate specific strings (use option --deobf to deobfuscate)':
    #TODO: regex to find several Chr*, not just one
        ('Chr', 'ChrB', 'ChrW', 'StrReverse', 'Xor'),
    #Chr: http://msdn.microsoft.com/en-us/library/office/gg264465%28v=office.15%29.aspx
    'May read or write registry keys':
    #sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
        ('RegOpenKeyExA', 'RegOpenKeyEx', 'RegCloseKey'),
    'May read registry keys':
    #sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
        ('RegQueryValueExA', 'RegQueryValueEx',
         'RegRead',  #with Wscript.Shell
        ),
    'May detect virtualization':
    # sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
        (r'SYSTEM\ControlSet001\Services\Disk\Enum', 'VIRTUAL', 'VMWARE', 'VBOX'),
    'May detect Anubis Sandbox':
    # sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
    # NOTES: this sample also checks App.EXEName but that seems to be a bug, it works in VB6 but not in VBA
    # ref: http://www.syssec-project.eu/m/page-media/3/disarm-raid11.pdf
        ('GetVolumeInformationA', 'GetVolumeInformation',  # with kernel32.dll
         '1824245000', r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductId',
         '76487-337-8429955-22614', 'andy', r'C:\exec\exec.exe', 'popupkiller'
         # note: removed 'sample' as it can trigger many false positives
        ),
    'May detect Sandboxie':
    # sample: https://malwr.com/analysis/M2NjZWNmMjA0YjVjNGVhYmJlZmFhNWY4NmQxZDllZTY/
    # ref: http://www.cplusplus.com/forum/windows/96874/
        ('SbieDll.dll', 'SandboxieControlWndClass'),
    'May detect Sunbelt Sandbox':
    # ref: http://www.cplusplus.com/forum/windows/96874/
        (r'C:\file.exe',),
    'May detect Norman Sandbox':
    # ref: http://www.cplusplus.com/forum/windows/96874/
        ('currentuser',),
    'May detect CW Sandbox':
    # ref: http://www.cplusplus.com/forum/windows/96874/
        ('Schmidti',),
    'May detect WinJail Sandbox':
    # ref: http://www.cplusplus.com/forum/windows/96874/
        ('Afx:400000:0',),
    'May attempt to disable VBA macro security and Protected View':
    # ref: http://blog.trendmicro.com/trendlabs-security-intelligence/qkg-filecoder-self-replicating-document-encrypting-ransomware/
    # ref: https://thehackernews.com/2017/11/ms-office-macro-malware.html
        ('AccessVBOM', 'VBAWarnings', 'ProtectedView', 'DisableAttachementsInPV', 'DisableInternetFilesInPV',
         'DisableUnsafeLocationsInPV', 'blockcontentexecutionfrominternet'),
    'May attempt to modify the VBA code (self-modification)':
        ('VBProject', 'VBComponents', 'CodeModule', 'AddFromString'),
    'May modify Excel 4 Macro formulas at runtime (XLM/XLF)':
        ('FORMULA.FILL',),
}

# Suspicious Keywords to be searched for directly as regex, without escaping
SUSPICIOUS_KEYWORDS_REGEX = {
    'May use Word Document Variables to store and hide data':
        (r'\.\s*Variables',),  # '.Variables' with optional whitespaces after the dot
                               # Vidar sample: https://app.any.run/tasks/897f28e7-3162-4b65-b268-2655543199d6/
    'May run a shellcode in memory':
        (r'EnumSystemLanguageGroupsW?', # Used by Hancitor in Oct 2016
         r'EnumDateFormats(?:W|(?:Ex){1,2})?', # see https://msdn.microsoft.com/en-us/library/windows/desktop/dd317810(v=vs.85).aspx
         ),
    'May run an executable file or a system command on a Mac (if combined with libc.dylib)':
        ('system', 'popen', r'exec[lv][ep]?'),
    'May run an executable file or a system command using Excel 4 Macros (XLM/XLF)':
        (r'(?<!Could contain following functions: )EXEC',),
    'Could contain a function that allows to run an executable file or a system command using Excel 4 Macros (XLM/XLF)':
        (r'Could contain following functions: EXEC',),
    'May call a DLL using Excel 4 Macros (XLM/XLF)':
        (r'(?<!Could contain following functions: )REGISTER',),
    'Could contain a function that allows to call a DLL using Excel 4 Macros (XLM/XLF)':
        (r'Could contain following functions: REGISTER',),
}

# Suspicious Keywords to be searched for directly as strings, without regex
SUSPICIOUS_KEYWORDS_NOREGEX = {
    'May use special characters such as backspace to obfuscate code when printed on the console':
        ('\b',),
}

# Regular Expression for a URL:
# http://en.wikipedia.org/wiki/Uniform_resource_locator
# http://www.w3.org/Addressing/URL/uri-spec.html
#TODO: also support username:password@server
#TODO: other protocols (file, gopher, wais, ...?)
SCHEME = r'\b(?:http|ftp)s?'
# see http://en.wikipedia.org/wiki/List_of_Internet_top-level_domains
TLD = r'(?:xn--[a-zA-Z0-9]{4,20}|[a-zA-Z]{2,20})'
DNS_NAME = r'(?:[a-zA-Z0-9\-\.]+\.' + TLD + ')'
#TODO: IPv6 - see https://www.debuggex.com/
# A literal numeric IPv6 address may be given, but must be enclosed in [ ] e.g. [db8:0cec::99:123a]
NUMBER_0_255 = r'(?:25[0-5]|2[0-4][0-9]|1[0-9]{2}|[1-9][0-9]|[0-9])'
IPv4 = r'(?:' + NUMBER_0_255 + r'\.){3}' + NUMBER_0_255
# IPv4 must come before the DNS name because it is more specific
SERVER = r'(?:' + IPv4 + '|' + DNS_NAME + ')'
PORT = r'(?:\:[0-9]{1,5})?'
SERVER_PORT = SERVER + PORT
URL_PATH = r'(?:/[a-zA-Z0-9\-\._\?\,\'/\\\+&%\$#\=~]*)?'  # [^\.\,\)\(\s"]
URL_RE = SCHEME + r'\://' + SERVER_PORT + URL_PATH
re_url = re.compile(URL_RE)

EXCLUDE_URLS_PATTERNS = ["http://schemas.openxmlformats.org/",
                         "http://schemas.microsoft.com/",
                         ]

# Patterns to be extracted (IP addresses, URLs, etc)
# From patterns.py in balbuzard
RE_PATTERNS = (
    ('URL', re.compile(URL_RE)),
    ('IPv4 address', re.compile(IPv4)),
    # TODO: add IPv6
    ('E-mail address', re.compile(r'(?i)\b[A-Z0-9._%+-]+@' + SERVER + '\b')),
    # ('Domain name', re.compile(r'(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)')),
    # Executable file name with known extensions (except .com which is present in many URLs, and .application):
    ("Executable file name", re.compile(
        r"(?i)\b\w+\.(EXE|PIF|GADGET|MSI|MSP|MSC|VBS|VBE|VB|JSE|JS|WSF|WSC|WSH|WS|BAT|CMD|DLL|SCR|HTA|CPL|CLASS|JAR|PS1XML|PS1|PS2XML|PS2|PSC1|PSC2|SCF|LNK|INF|REG)\b")),
    # Sources: http://www.howtogeek.com/137270/50-file-extensions-that-are-potentially-dangerous-on-windows/
    # TODO: https://support.office.com/en-us/article/Blocked-attachments-in-Outlook-3811cddc-17c3-4279-a30c-060ba0207372#__attachment_file_types
    # TODO: add win & unix file paths
    #('Hex string', re.compile(r'(?:[0-9A-Fa-f]{2}){4,}')),
)

# regex to detect strings encoded in hexadecimal
re_hex_string = re.compile(r'(?:[0-9A-Fa-f]{2}){4,}')

# regex to detect strings encoded in base64
#re_base64_string = re.compile(r'"(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?"')
# better version from balbuzard, less false positives:
# (plain version without double quotes, used also below in quoted_base64_string)
BASE64_RE = r'(?:[A-Za-z0-9+/]{4}){1,}(?:[A-Za-z0-9+/]{2}[AEIMQUYcgkosw048]=|[A-Za-z0-9+/][AQgw]==)?'
re_base64_string = re.compile('"' + BASE64_RE + '"')
# white list of common strings matching the base64 regex, but which are not base64 strings (all lowercase):
BASE64_WHITELIST = set(['thisdocument', 'thisworkbook', 'test', 'temp', 'http', 'open', 'exit', 'kernel32',
                        'virtualalloc', 'createthread'])

# regex to detect strings encoded with a specific Dridex algorithm
# (see https://github.com/JamesHabben/MalwareStuff)
re_dridex_string = re.compile(r'"[0-9A-Za-z]{20,}"')
# regex to check that it is not just a hex string:
re_nothex_check = re.compile(r'[G-Zg-z]')

# regex to extract printable strings (at least 5 chars) from VBA Forms:
# (must be bytes for Python 3)
re_printable_string = re.compile(b'[\\t\\r\\n\\x20-\\xFF]{5,}')


# === PARTIAL VBA GRAMMAR ====================================================

# REFERENCES:
# - [MS-VBAL]: VBA Language Specification
#   https://msdn.microsoft.com/en-us/library/dd361851.aspx
# - pyparsing: http://pyparsing.wikispaces.com/

# TODO: set whitespaces according to VBA
# TODO: merge extended lines before parsing

# Enable PackRat for better performance:
# (see https://pythonhosted.org/pyparsing/pyparsing.ParserElement-class.html#enablePackrat)
ParserElement.enablePackrat()

# VBA identifier chars (from MS-VBAL 3.3.5)
vba_identifier_chars = alphanums + '_'

class VbaExpressionString(str):
    """
    Class identical to str, used to distinguish plain strings from strings
    obfuscated using VBA expressions (Chr, StrReverse, etc)
    Usage: each VBA expression parse action should convert strings to
    VbaExpressionString.
    Then isinstance(s, VbaExpressionString) is True only for VBA expressions.
     (see detect_vba_strings)
    """
    # TODO: use Unicode everywhere instead of str
    pass


# --- NUMBER TOKENS ----------------------------------------------------------

# 3.3.2 Number Tokens
# INTEGER = integer-literal ["%" / "&" / "^"]
# integer-literal = decimal-literal / octal-literal / hex-literal
# decimal-literal = 1*decimal-digit
# octal-literal = "&" [%x004F / %x006F] 1*octal-digit
# ; & or &o or &O
# hex-literal = "&" (%x0048 / %x0068) 1*hex-digit
# ; &h or &H
# octal-digit = "0" / "1" / "2" / "3" / "4" / "5" / "6" / "7"
# decimal-digit = octal-digit / "8" / "9"
# hex-digit = decimal-digit / %x0041-0046 / %x0061-0066 ;A-F / a-f

# NOTE: here Combine() is required to avoid spaces between elements
# NOTE: here WordStart is necessary to avoid matching a number preceded by
#       letters or underscore (e.g. "VBT1" or "ABC_34"), when using scanString
decimal_literal = Combine(Optional('-') + WordStart(vba_identifier_chars) + Word(nums)
                          + Suppress(Optional(Word('%&^', exact=1))))
decimal_literal.setParseAction(lambda t: int(t[0]))

octal_literal = Combine(Suppress(Literal('&') + Optional((CaselessLiteral('o')))) + Word(srange('[0-7]'))
                + Suppress(Optional(Word('%&^', exact=1))))
octal_literal.setParseAction(lambda t: int(t[0], base=8))

hex_literal = Combine(Suppress(CaselessLiteral('&h')) + Word(srange('[0-9a-fA-F]'))
                + Suppress(Optional(Word('%&^', exact=1))))
hex_literal.setParseAction(lambda t: int(t[0], base=16))

integer = decimal_literal | octal_literal | hex_literal


# --- QUOTED STRINGS ---------------------------------------------------------

# 3.3.4 String Tokens
# STRING = double-quote *string-character (double-quote / line-continuation / LINE-END)
# double-quote = %x0022 ; "
# string-character = NO-LINE-CONTINUATION ((double-quote double-quote) termination-character)

quoted_string = QuotedString('"', escQuote='""')
quoted_string.setParseAction(lambda t: str(t[0]))


#--- VBA Expressions ---------------------------------------------------------

# See MS-VBAL 5.6 Expressions

# need to pre-declare using Forward() because it is recursive
# VBA string expression and integer expression
vba_expr_str = Forward()
vba_expr_int = Forward()

# --- CHR --------------------------------------------------------------------

# MS-VBAL 6.1.2.11.1.4 Chr / Chr$
# Function Chr(CharCode As Long) As Variant
# Function Chr$(CharCode As Long) As String
# Parameter Description
# CharCode Long whose value is a code point.
# Returns a String data value consisting of a single character containing the character whose code
# point is the data value of the argument.
# - If the argument is not in the range 0 to 255, Error Number 5 ("Invalid procedure call or
# argument") is raised unless the implementation supports a character set with a larger code point
# range.
# - If the argument value is in the range of 0 to 127, it is interpreted as a 7-bit ASCII code point.
# - If the argument value is in the range of 128 to 255, the code point interpretation of the value is
# implementation defined.
# - Chr$ has the same runtime semantics as Chr, however the declared type of its function result is
# String rather than Variant.

# 6.1.2.11.1.5 ChrB / ChrB$
# Function ChrB(CharCode As Long) As Variant
# Function ChrB$(CharCode As Long) As String
# CharCode Long whose value is a code point.
# Returns a String data value consisting of a single byte character whose code point value is the
# data value of the argument.
# - If the argument is not in the range 0 to 255, Error Number 6 ("Overflow") is raised.
# - ChrB$ has the same runtime semantics as ChrB however the declared type of its function result
# is String rather than Variant.
# - Note: the ChrB function is used with byte data contained in a String. Instead of returning a
# character, which may be one or two bytes, ChrB always returns a single byte. The ChrW function
# returns a String containing the Unicode character except on platforms where Unicode is not
# supported, in which case, the behavior is identical to the Chr function.

# 6.1.2.11.1.6 ChrW/ ChrW$
# Function ChrW(CharCode As Long) As Variant
# Function ChrW$(CharCode As Long) As String
# CharCode Long whose value is a code point.
# Returns a String data value consisting of a single character containing the character whose code
# point is the data value of the argument.
# - If the argument is not in the range -32,767 to 65,535 then Error Number 5 ("Invalid procedure
# call or argument") is raised.
# - If the argument is a negative value it is treated as if it was the value: CharCode + 65,536.
# - If the implemented uses 16-bit Unicode code points argument, data value is interpreted as a 16-
# bit Unicode code point.
# - If the implementation does not support Unicode, ChrW has the same semantics as Chr.
# - ChrW$ has the same runtime semantics as ChrW, however the declared type of its function result
# is String rather than Variant.

# Chr, Chr$, ChrB, ChrW(int) => char
vba_chr = Suppress(
            Combine(WordStart(vba_identifier_chars) + CaselessLiteral('Chr')
            + Optional(CaselessLiteral('B') | CaselessLiteral('W')) + Optional('$'))
            + '(') + vba_expr_int + Suppress(')')

def vba_chr_tostr(t):
    try:
        i = t[0]
        if i>=0 and i<=255:
            # normal, non-unicode character:
            # TODO: check if it needs to be converted to bytes for Python 3
            return VbaExpressionString(chr(i))
        else:
            # unicode character
            # Note: this distinction is only needed for Python 2
            # pylint: disable-next=possibly-used-before-assignment
            return VbaExpressionString(unichr(i).encode('utf-8', 'backslashreplace'))
    except ValueError:
        log.exception('ERROR: incorrect parameter value for chr(): %r' % i)
        return VbaExpressionString('Chr(%r)' % i)

vba_chr.setParseAction(vba_chr_tostr)


# --- ASC --------------------------------------------------------------------

# Asc(char) => int
#TODO: see MS-VBAL 6.1.2.11.1.1 page 240 => AscB, AscW
vba_asc = Suppress(CaselessKeyword('Asc') + '(') + vba_expr_str + Suppress(')')
vba_asc.setParseAction(lambda t: ord(t[0]))


# --- VAL --------------------------------------------------------------------

# Val(string) => int
# TODO: make sure the behavior of VBA's val is fully covered
vba_val = Suppress(CaselessKeyword('Val') + '(') + vba_expr_str + Suppress(')')
vba_val.setParseAction(lambda t: int(t[0].strip()))


# --- StrReverse() --------------------------------------------------------------------

# StrReverse(string) => string
strReverse = Suppress(CaselessKeyword('StrReverse') + '(') + vba_expr_str + Suppress(')')
strReverse.setParseAction(lambda t: VbaExpressionString(str(t[0])[::-1]))


# --- ENVIRON() --------------------------------------------------------------------

# Environ("name") => just translated to "%name%", that is enough for malware analysis
environ = Suppress(CaselessKeyword('Environ') + '(') + vba_expr_str + Suppress(')')
environ.setParseAction(lambda t: VbaExpressionString('%%%s%%' % t[0]))


# --- IDENTIFIER -------------------------------------------------------------

#TODO: see MS-VBAL 3.3.5 page 33
# 3.3.5 Identifier Tokens
# Latin-identifier = first-Latin-identifier-character *subsequent-Latin-identifier-character
# first-Latin-identifier-character = (%x0041-005A / %x0061-007A) ; A-Z / a-z
# subsequent-Latin-identifier-character = first-Latin-identifier-character / DIGIT / %x5F ; underscore
latin_identifier = Word(initChars=alphas, bodyChars=alphanums + '_')

# --- HEX FUNCTION -----------------------------------------------------------

# match any custom function name with a hex string as argument:
# TODO: accept vba_expr_str_item as argument, check if it is a hex or base64 string at runtime

# quoted string of at least two hexadecimal numbers of two digits:
quoted_hex_string = Suppress('"') + Combine(Word(hexnums, exact=2) * (2, None)) + Suppress('"')
quoted_hex_string.setParseAction(lambda t: str(t[0]))

hex_function_call = Suppress(latin_identifier) + Suppress('(') + \
                    quoted_hex_string('hex_string') + Suppress(')')
hex_function_call.setParseAction(lambda t: VbaExpressionString(binascii.a2b_hex(t.hex_string)))


# --- BASE64 FUNCTION -----------------------------------------------------------

# match any custom function name with a Base64 string as argument:
# TODO: accept vba_expr_str_item as argument, check if it is a hex or base64 string at runtime

# quoted string of at least two hexadecimal numbers of two digits:
quoted_base64_string = Suppress('"') + Regex(BASE64_RE) + Suppress('"')
quoted_base64_string.setParseAction(lambda t: str(t[0]))

base64_function_call = Suppress(latin_identifier) + Suppress('(') + \
                    quoted_base64_string('base64_string') + Suppress(')')
base64_function_call.setParseAction(lambda t: VbaExpressionString(binascii.a2b_base64(t.base64_string)))


# ---STRING EXPRESSION -------------------------------------------------------

def concat_strings_list(tokens):
    """
    parse action to concatenate strings in a VBA expression with operators '+' or '&'
    """
    # extract argument from the tokens:
    # expected to be a tuple containing a list of strings such as [a,'&',b,'&',c,...]
    strings = tokens[0][::2]
    return VbaExpressionString(''.join(strings))


vba_expr_str_item = (vba_chr | strReverse | environ | quoted_string | hex_function_call | base64_function_call)

vba_expr_str <<= infixNotation(vba_expr_str_item,
    [
        ("+", 2, opAssoc.LEFT, concat_strings_list),
        ("&", 2, opAssoc.LEFT, concat_strings_list),
    ])


# --- INTEGER EXPRESSION -------------------------------------------------------

def sum_ints_list(tokens):
    """
    parse action to sum integers in a VBA expression with operator '+'
    """
    # extract argument from the tokens:
    # expected to be a tuple containing a list of integers such as [a,'&',b,'&',c,...]
    integers = tokens[0][::2]
    return sum(integers)


def subtract_ints_list(tokens):
    """
    parse action to subtract integers in a VBA expression with operator '-'
    """
    # extract argument from the tokens:
    # expected to be a tuple containing a list of integers such as [a,'&',b,'&',c,...]
    integers = tokens[0][::2]
    # pylint: disable-next=possibly-used-before-assignment
    return reduce(lambda x,y:x-y, integers)


def multiply_ints_list(tokens):
    """
    parse action to multiply integers in a VBA expression with operator '*'
    """
    # extract argument from the tokens:
    # expected to be a tuple containing a list of integers such as [a,'&',b,'&',c,...]
    integers = tokens[0][::2]
    return reduce(lambda x,y:x*y, integers)


def divide_ints_list(tokens):
    """
    parse action to divide integers in a VBA expression with operator '/'
    """
    # extract argument from the tokens:
    # expected to be a tuple containing a list of integers such as [a,'&',b,'&',c,...]
    integers = tokens[0][::2]
    return reduce(lambda x,y:x/y, integers)


vba_expr_int_item = (vba_asc | vba_val | integer)

# operators associativity:
# https://en.wikipedia.org/wiki/Operator_associativity

vba_expr_int <<= infixNotation(vba_expr_int_item,
    [
        ("*", 2, opAssoc.LEFT, multiply_ints_list),
        ("/", 2, opAssoc.LEFT, divide_ints_list),
        ("-", 2, opAssoc.LEFT, subtract_ints_list),
        ("+", 2, opAssoc.LEFT, sum_ints_list),
    ])


# see detect_vba_strings for the deobfuscation code using this grammar

# === MSO/ActiveMime files parsing ===========================================

def is_mso_file(data):
    """
    Check if the provided data is the content of a MSO/ActiveMime file, such as
    the ones created by Outlook in some cases, or Word/Excel when saving a
    file with the MHTML format or the Word 2003 XML format.
    This function only checks the ActiveMime magic at the beginning of data.
    :param data: bytes string, MSO/ActiveMime file content
    :return: bool, True if the file is MSO, False otherwise
    """
    return data.startswith(MSO_ACTIVEMIME_HEADER)


# regex to find zlib block headers, starting with byte 0x78 = 'x'
re_zlib_header = re.compile(r'x')


def mso_file_extract(data):
    """
    Extract the data stored into a MSO/ActiveMime file, such as
    the ones created by Outlook in some cases, or Word/Excel when saving a
    file with the MHTML format or the Word 2003 XML format.

    :param data: bytes string, MSO/ActiveMime file content
    :return: bytes string, extracted data (uncompressed)

    raise a MsoExtractionError if the data cannot be extracted
    """
    # check the magic:
    assert is_mso_file(data)

    # In all the samples seen so far, Word always uses an offset of 0x32,
    # and Excel 0x22A. But we read the offset from the header to be more
    # generic.
    offsets = [0x32, 0x22A]

    # First, attempt to get the compressed data offset from the header
    # According to my tests, it should be an unsigned 16 bits integer,
    # at offset 0x1E (little endian) + add 46:
    try:
        offset = struct.unpack_from('<H', data, offset=0x1E)[0] + 46
        log.debug('Parsing MSO file: data offset = 0x%X' % offset)
        offsets.insert(0, offset)  # insert at beginning of offsets
    except struct.error as exc:
        log.info('Unable to parse MSO/ActiveMime file header (%s)' % exc)
        log.debug('Trace:', exc_info=True)
        raise MsoExtractionError('Unable to parse MSO/ActiveMime file header')
    # now try offsets
    for start in offsets:
        try:
            log.debug('Attempting zlib decompression from MSO file offset 0x%X' % start)
            extracted_data = zlib.decompress(data[start:])
            return extracted_data
        except zlib.error as exc:
            log.info('zlib decompression failed for offset %s (%s)'
                     % (start, exc))
            log.debug('Trace:', exc_info=True)
    # None of the guessed offsets worked, let's try brute-forcing by looking
    # for potential zlib-compressed blocks starting with 0x78:
    log.debug('Looking for potential zlib-compressed blocks in MSO file')
    for match in re_zlib_header.finditer(data):
        start = match.start()
        try:
            log.debug('Attempting zlib decompression from MSO file offset 0x%X' % start)
            extracted_data = zlib.decompress(data[start:])
            return extracted_data
        except zlib.error as exc:
            log.info('zlib decompression failed (%s)' % exc)
            log.debug('Trace:', exc_info=True)
    raise MsoExtractionError('Unable to decompress data from a MSO/ActiveMime file')


#--- FUNCTIONS ----------------------------------------------------------------

# set of printable characters, for is_printable
_PRINTABLE_SET = set(string.printable)

def is_printable(s):
    """
    returns True if string s only contains printable ASCII characters
    (i.e. contained in string.printable)
    This is similar to Python 3's str.isprintable, for Python 2.x.
    :param s: str
    :return: bool
    """
    # inspired from http://stackoverflow.com/questions/3636928/test-if-a-python-string-is-printable
    # check if the set of chars from s is contained into the set of printable chars:
    return set(s).issubset(_PRINTABLE_SET)


def copytoken_help(decompressed_current, decompressed_chunk_start):
    """
    compute bit masks to decode a CopyToken according to MS-OVBA 2.4.1.3.19.1 CopyToken Help

    decompressed_current: number of decompressed bytes so far, i.e. len(decompressed_container)
    decompressed_chunk_start: offset of the current chunk in the decompressed container
    return length_mask, offset_mask, bit_count, maximum_length
    """
    difference = decompressed_current - decompressed_chunk_start
    bit_count = int(math.ceil(math.log(difference, 2)))
    bit_count = max([bit_count, 4])
    length_mask = 0xFFFF >> bit_count
    offset_mask = ~length_mask
    maximum_length = (0xFFFF >> bit_count) + 3
    return length_mask, offset_mask, bit_count, maximum_length


def decompress_stream(compressed_container):
    """
    Decompress a stream according to MS-OVBA section 2.4.1

    :param compressed_container bytearray: bytearray or bytes compressed according to the MS-OVBA 2.4.1.3.6 Compression algorithm
    :return: the decompressed container as a bytes string
    :rtype: bytes
    """
    # 2.4.1.2 State Variables

    # The following state is maintained for the CompressedContainer (section 2.4.1.1.1):
    # CompressedRecordEnd: The location of the byte after the last byte in the CompressedContainer (section 2.4.1.1.1).
    # CompressedCurrent: The location of the next byte in the CompressedContainer (section 2.4.1.1.1) to be read by
    #                    decompression or to be written by compression.

    # The following state is maintained for the current CompressedChunk (section 2.4.1.1.4):
    # CompressedChunkStart: The location of the first byte of the CompressedChunk (section 2.4.1.1.4) within the
    #                       CompressedContainer (section 2.4.1.1.1).

    # The following state is maintained for a DecompressedBuffer (section 2.4.1.1.2):
    # DecompressedCurrent: The location of the next byte in the DecompressedBuffer (section 2.4.1.1.2) to be written by
    #                      decompression or to be read by compression.
    # DecompressedBufferEnd: The location of the byte after the last byte in the DecompressedBuffer (section 2.4.1.1.2).

    # The following state is maintained for the current DecompressedChunk (section 2.4.1.1.3):
    # DecompressedChunkStart: The location of the first byte of the DecompressedChunk (section 2.4.1.1.3) within the
    #                         DecompressedBuffer (section 2.4.1.1.2).

    # Check the input is a bytearray, otherwise convert it (assuming it's bytes):
    if not isinstance(compressed_container, bytearray):
        compressed_container = bytearray(compressed_container)
        # raise TypeError('decompress_stream requires a bytearray as input')
    log.debug('decompress_stream: compressed size = {} bytes'.format(len(compressed_container)))
    decompressed_container = bytearray()  # result
    compressed_current = 0

    sig_byte = compressed_container[compressed_current]
    if sig_byte != 0x01:
        raise ValueError('invalid signature byte {0:02X}'.format(sig_byte))

    compressed_current += 1

    #NOTE: the definition of CompressedRecordEnd is ambiguous. Here we assume that
    # CompressedRecordEnd = len(compressed_container)
    while compressed_current < len(compressed_container):
        # 2.4.1.1.5
        compressed_chunk_start = compressed_current
        # chunk header = first 16 bits
        compressed_chunk_header = \
            struct.unpack("<H", compressed_container[compressed_chunk_start:compressed_chunk_start + 2])[0]
        # chunk size = 12 first bits of header + 3
        chunk_size = (compressed_chunk_header & 0x0FFF) + 3
        # chunk signature = 3 next bits - should always be 0b011
        chunk_signature = (compressed_chunk_header >> 12) & 0x07
        if chunk_signature != 0b011:
            raise ValueError('Invalid CompressedChunkSignature in VBA compressed stream')
        # chunk flag = next bit - 1 == compressed, 0 == uncompressed
        chunk_flag = (compressed_chunk_header >> 15) & 0x01
        log.debug("chunk size = {}, offset = {}, compressed flag = {}".format(chunk_size, compressed_chunk_start, chunk_flag))

        #MS-OVBA 2.4.1.3.12: the maximum size of a chunk including its header is 4098 bytes (header 2 + data 4096)
        # The minimum size is 3 bytes
        # NOTE: there seems to be a typo in MS-OVBA, the check should be with 4098, not 4095 (which is the max value
        # in chunk header before adding 3.
        # Also the first test is not useful since a 12 bits value cannot be larger than 4095.
        if chunk_flag == 1 and chunk_size > 4098:
            raise ValueError('CompressedChunkSize=%d > 4098 but CompressedChunkFlag == 1' % chunk_size)
        if chunk_flag == 0 and chunk_size != 4098:
            raise ValueError('CompressedChunkSize=%d != 4098 but CompressedChunkFlag == 0' % chunk_size)

        # check if chunk_size goes beyond the compressed data, instead of silently cutting it:
        #TODO: raise an exception?
        if compressed_chunk_start + chunk_size > len(compressed_container):
            log.warning('Chunk size is larger than remaining compressed data')
        compressed_end = min([len(compressed_container), compressed_chunk_start + chunk_size])
        # read after chunk header:
        compressed_current = compressed_chunk_start + 2

        if chunk_flag == 0:
            # MS-OVBA 2.4.1.3.3 Decompressing a RawChunk
            # uncompressed chunk: read the next 4096 bytes as-is
            #TODO: check if there are at least 4096 bytes left
            decompressed_container.extend(compressed_container[compressed_current:compressed_current + 4096])
            compressed_current += 4096
        else:
            # MS-OVBA 2.4.1.3.2 Decompressing a CompressedChunk
            # compressed chunk
            decompressed_chunk_start = len(decompressed_container)
            while compressed_current < compressed_end:
                # MS-OVBA 2.4.1.3.4 Decompressing a TokenSequence
                # log.debug('compressed_current = %d / compressed_end = %d' % (compressed_current, compressed_end))
                # FlagByte: 8 bits indicating if the following 8 tokens are either literal (1 byte of plain text) or
                # copy tokens (reference to a previous literal token)
                flag_byte = compressed_container[compressed_current]
                compressed_current += 1
                # pylint: disable-next=possibly-used-before-assignment
                for bit_index in xrange(0, 8):
                    # log.debug('bit_index=%d / compressed_current=%d / compressed_end=%d' % (bit_index, compressed_current, compressed_end))
                    if compressed_current >= compressed_end:
                        break
                    # MS-OVBA 2.4.1.3.5 Decompressing a Token
                    # MS-OVBA 2.4.1.3.17 Extract FlagBit
                    flag_bit = (flag_byte >> bit_index) & 1
                    #log.debug('bit_index=%d: flag_bit=%d' % (bit_index, flag_bit))
                    if flag_bit == 0:  # LiteralToken
                        # copy one byte directly to output
                        decompressed_container.extend([compressed_container[compressed_current]])
                        compressed_current += 1
                    else:  # CopyToken
                        # MS-OVBA 2.4.1.3.19.2 Unpack CopyToken
                        copy_token = \
                            struct.unpack("<H", compressed_container[compressed_current:compressed_current + 2])[0]
                        #TODO: check this
                        length_mask, offset_mask, bit_count, _ = copytoken_help(
                            len(decompressed_container), decompressed_chunk_start)
                        length = (copy_token & length_mask) + 3
                        temp1 = copy_token & offset_mask
                        temp2 = 16 - bit_count
                        offset = (temp1 >> temp2) + 1
                        #log.debug('offset=%d length=%d' % (offset, length))
                        copy_source = len(decompressed_container) - offset
                        for index in xrange(copy_source, copy_source + length):
                            decompressed_container.extend([decompressed_container[index]])
                        compressed_current += 2
    return bytes(decompressed_container)


class VBA_Module(object):
    """
    Class to parse a VBA module from an OLE file, and to store all the corresponding
    metadata and VBA source code.
    """

    def __init__(self, project, dir_stream, module_index):
        """
        Parse a VBA Module record from the dir stream of a VBA project.
        Reference: MS-OVBA 2.3.4.2.3.2 MODULE Record

        :param VBA_Project project: VBA_Project, corresponding VBA project
        :param olefile.OleStream dir_stream: olefile.OleStream, file object containing the module record
        :param int module_index: int, index of the module in the VBA project list
        """
        #: reference to the VBA project for later use (VBA_Project)
        self.project = project
        #: VBA module name (unicode str)
        self.name = None
        #: VBA module name as a native str (utf8 bytes on py2, str on py3)
        self.name_str = None
        #: VBA module name, unicode copy (unicode str)
        self._name_unicode = None
        #: Stream name containing the VBA module (unicode str)
        self.streamname = None
        #: Stream name containing the VBA module as a native str (utf8 bytes on py2, str on py3)
        self.streamname_str = None
        self._streamname_unicode = None
        self.docstring = None
        self._docstring_unicode = None
        self.textoffset = None
        self.type = None
        self.readonly = False
        self.private = False
        #: VBA source code in bytes format, using the original code page from the VBA project
        self.code_raw = None
        #: VBA source code in unicode format (unicode for Python2, str for Python 3)
        self.code = None
        #: VBA source code in native str format (str encoded with UTF-8 for Python 2, str for Python 3)
        self.code_str = None
        #: VBA module file name including an extension based on the module type such as bas, cls, frm (unicode str)
        self.filename = None
        #: VBA module file name in native str format (str)
        self.filename_str = None
        self.code_path = None
        try:
            # 2.3.4.2.3.2.1 MODULENAME Record
            # Specifies a VBA identifier as the name of the containing MODULE Record
            _id = struct.unpack("<H", dir_stream.read(2))[0]
            project.check_value('MODULENAME_Id', 0x0019, _id)
            size = struct.unpack("<L", dir_stream.read(4))[0]
            modulename_bytes = dir_stream.read(size)
            # Module name always stored as Unicode:
            self.name = project.decode_bytes(modulename_bytes)
            self.name_str = unicode2str(self.name)
            # account for optional sections
            # TODO: shouldn't this be a loop? (check MS-OVBA)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x0047:
                # 2.3.4.2.3.2.2 MODULENAMEUNICODE Record
                # Specifies a VBA identifier as the name of the containing MODULE Record (section 2.3.4.2.3.2).
                # MUST contain the UTF-16 encoding of MODULENAME Record
                size = struct.unpack("<L", dir_stream.read(4))[0]
                self._name_unicode = dir_stream.read(size).decode('UTF-16LE', 'replace')
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x001A:
                # 2.3.4.2.3.2.3 MODULESTREAMNAME Record
                # Specifies the stream name of the ModuleStream (section 2.3.4.3) in the VBA Storage (section 2.3.4)
                # corresponding to the containing MODULE Record
                size = struct.unpack("<L", dir_stream.read(4))[0]
                streamname_bytes = dir_stream.read(size)
                # Store it as Unicode:
                self.streamname = project.decode_bytes(streamname_bytes)
                self.streamname_str = unicode2str(self.streamname)
                reserved = struct.unpack("<H", dir_stream.read(2))[0]
                project.check_value('MODULESTREAMNAME_Reserved', 0x0032, reserved)
                size = struct.unpack("<L", dir_stream.read(4))[0]
                self._streamname_unicode = dir_stream.read(size).decode('UTF-16LE', 'replace')
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x001C:
                # 2.3.4.2.3.2.4 MODULEDOCSTRING Record
                # Specifies the description for the containing MODULE Record
                size = struct.unpack("<L", dir_stream.read(4))[0]
                docstring_bytes = dir_stream.read(size)
                self.docstring = project.decode_bytes(docstring_bytes)
                reserved = struct.unpack("<H", dir_stream.read(2))[0]
                project.check_value('MODULEDOCSTRING_Reserved', 0x0048, reserved)
                size = struct.unpack("<L", dir_stream.read(4))[0]
                self._docstring_unicode = dir_stream.read(size)
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x0031:
                # 2.3.4.2.3.2.5 MODULEOFFSET Record
                # Specifies the location of the source code within the ModuleStream (section 2.3.4.3)
                # that corresponds to the containing MODULE Record
                size = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULEOFFSET_Size', 0x0004, size)
                self.textoffset = struct.unpack("<L", dir_stream.read(4))[0]
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x001E:
                # 2.3.4.2.3.2.6 MODULEHELPCONTEXT Record
                # Specifies the Help topic identifier for the containing MODULE Record
                modulehelpcontext_size = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULEHELPCONTEXT_Size', 0x0004, modulehelpcontext_size)
                # HelpContext (4 bytes): An unsigned integer that specifies the Help topic identifier
                # in the Help file specified by PROJECTHELPFILEPATH Record
                helpcontext = struct.unpack("<L", dir_stream.read(4))[0]
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x002C:
                # 2.3.4.2.3.2.7 MODULECOOKIE Record
                # Specifies ignored data.
                size = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULECOOKIE_Size', 0x0002, size)
                cookie = struct.unpack("<H", dir_stream.read(2))[0]
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x0021 or section_id == 0x0022:
                # 2.3.4.2.3.2.8 MODULETYPE Record
                # Specifies whether the containing MODULE Record (section 2.3.4.2.3.2) is a procedural module,
                # document module, class module, or designer module.
                # Id (2 bytes): An unsigned integer that specifies the identifier for this record.
                # MUST be 0x0021 when the containing MODULE Record (section 2.3.4.2.3.2) is a procedural module.
                # MUST be 0x0022 when the containing MODULE Record (section 2.3.4.2.3.2) is a document module,
                # class module, or designer module.
                self.type = section_id
                reserved = struct.unpack("<L", dir_stream.read(4))[0]
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x0025:
                # 2.3.4.2.3.2.9 MODULEREADONLY Record
                # Specifies that the containing MODULE Record (section 2.3.4.2.3.2) is read-only.
                self.readonly = True
                reserved = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULEREADONLY_Reserved', 0x0000, reserved)
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x0028:
                # 2.3.4.2.3.2.10 MODULEPRIVATE Record
                # Specifies that the containing MODULE Record (section 2.3.4.2.3.2) is only usable from within
                # the current VBA project.
                self.private = True
                reserved = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULEPRIVATE_Reserved', 0x0000, reserved)
                section_id = struct.unpack("<H", dir_stream.read(2))[0]
            if section_id == 0x002B:  # TERMINATOR
                # Terminator (2 bytes): An unsigned integer that specifies the end of this record. MUST be 0x002B.
                # Reserved (4 bytes): MUST be 0x00000000. MUST be ignored.
                reserved = struct.unpack("<L", dir_stream.read(4))[0]
                project.check_value('MODULE_Reserved', 0x0000, reserved)
                section_id = None
            if section_id != None:
                log.warning('unknown or invalid module section id {0:04X}'.format(section_id))
        
            log.debug("Module Name = {0}".format(self.name_str))
            # log.debug("Module Name Unicode = {0}".format(self._name_unicode))
            log.debug("Stream Name = {0}".format(self.streamname_str))
            # log.debug("Stream Name Unicode = {0}".format(self._streamname_unicode))
            log.debug("TextOffset = {0}".format(self.textoffset))
        
            code_data = None
            # let's try the different names we have, just in case some are missing:
            try_names = (self.streamname, self._streamname_unicode, self.name, self._name_unicode)
            for stream_name in try_names:
                # TODO: if olefile._find were less private, could replace this
                #        try-except with calls to it
                if stream_name is not None:
                    try:
                        self.code_path = project.vba_root + u'VBA/' + stream_name
                        log.debug('opening VBA code stream %s' % self.code_path)
                        code_data = project.ole.openstream(self.code_path).read()
                        break
                    except IOError as ioe:
                        log.debug('failed to open stream VBA/%r (%r), try other name'
                                  % (stream_name, ioe))
        
            if code_data is None:
                log.info("Could not open stream %d of %d ('VBA/' + one of %r)!"
                         % (module_index, project.modules_count,
                            '/'.join("'" + stream_name + "'"
                                     for stream_name in try_names)))
                if project.relaxed:
                    return  # ... continue with next submodule
                else:
                    raise SubstreamOpenError('[BASE]', 'VBA/' + self.name)
        
            log.debug("length of code_data = {0}".format(len(code_data)))
            log.debug("offset of code_data = {0}".format(self.textoffset))
            code_data = code_data[self.textoffset:]
            if len(code_data) > 0:
                code_data = decompress_stream(bytearray(code_data))
                # store the raw code encoded as bytes with the project's code page:
                self.code_raw = code_data
                # decode it to unicode:
                self.code = project.decode_bytes(code_data)
                # also store a native str version:
                self.code_str = unicode2str(self.code)
                # case-insensitive search in the code_modules dict to find the file extension:
                filext = self.project.module_ext.get(self.name.lower(), 'vba')
                self.filename = u'{0}.{1}'.format(self.name, filext)
                self.filename_str = unicode2str(self.filename)
                log.debug('extracted file {0}'.format(self.filename_str))
            else:
                log.warning("module stream {0} has code data length 0".format(self.streamname_str))
        except (UnexpectedDataError, SubstreamOpenError):
            raise
        except Exception as exc:
            log.info('Error parsing module {0} of {1}:'
                     .format(module_index, project.modules_count),
                     exc_info=True)
            # TODO: here if we don't raise the exception it causes other issues because the module
            #       is not properly initialised (e.g. self.code_str=None causing issue #629)
            raise
            # if not project.relaxed:
            #     raise


class VBA_Project(object):
    """
    Class to parse a VBA project from an OLE file, and to store all the corresponding
    metadata and VBA modules.
    """

    def __init__(self, ole, vba_root, project_path, dir_path, relaxed=True):
        """
        Extract VBA macros from an OleFileIO object.

        :param vba_root: path to the VBA root storage, containing the VBA storage and the PROJECT stream
        :param project_path: path to the PROJECT stream
        :param relaxed: If True, only create info/debug log entry if data is not as expected
                        (e.g. opening substream fails); if False, raise an error in this case
        """
        self.ole = ole
        self.vba_root = vba_root
        self. project_path = project_path
        self.dir_path = dir_path
        self.relaxed = relaxed
        #: VBA modules contained in the project (list of VBA_Module objects)
        self.modules = []
        #: file extension for each VBA module
        self.module_ext = {}
        log.debug('Parsing the dir stream from %r' % dir_path)
        # read data from dir stream (compressed)
        dir_compressed = ole.openstream(dir_path).read()
        # decompress it:
        dir_stream = BytesIO(decompress_stream(bytearray(dir_compressed)))
        # store reference for later use:
        self.dir_stream = dir_stream

        # reference: MS-VBAL 2.3.4.2 dir Stream: Version Independent Project Information

        # PROJECTSYSKIND Record
        # Specifies the platform for which the VBA project is created.
        projectsyskind_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTSYSKIND_Id', 0x0001, projectsyskind_id)
        projectsyskind_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTSYSKIND_Size', 0x0004, projectsyskind_size)
        self.syskind = struct.unpack("<L", dir_stream.read(4))[0]
        SYSKIND_NAME = {
            0x00: "16-bit Windows",
            0x01: "32-bit Windows",
            0x02: "Macintosh",
            0x03: "64-bit Windows"
        }
        self.syskind_name = SYSKIND_NAME.get(self.syskind, 'Unknown')
        log.debug("PROJECTSYSKIND_SysKind: %d - %s" % (self.syskind, self.syskind_name))
        if self.syskind not in SYSKIND_NAME:
            log.error("invalid PROJECTSYSKIND_SysKind {0:04X}".format(self.syskind))

        # PROJECTLCID Record or PROJECTCOMPATVERSION Record
        project_id = struct.unpack("<H", dir_stream.read(2))[0]
        if project_id == 0x004A:
            # PROJECTCOMPATVERSION Record
            # Specifies the VBA project's compat version.
            projectcompatversion_id = project_id
            self.check_value('PROJETCOMPATVERSION_Id', 0x004A, projectcompatversion_id)
            projectcompatversion_size = struct.unpack("<L", dir_stream.read(4))[0]
            self.check_value('PROJECTCOMPATVERSION_Size', 0x0004, projectcompatversion_size)
            projectcompatversion_compatversion = struct.unpack("<L", dir_stream.read(4))[0]
            # compat version: A 32-bit number that identifies the Office Model version used by a VBA project.
            log.debug("compat version: {compat_version}".format(compat_version=projectcompatversion_compatversion))

            # PROJECTLCID Record
            project_id = struct.unpack("<H", dir_stream.read(2))[0]

        projectlcid_id = project_id

        # Specifies the VBA project's LCID.
        self.check_value('PROJECTLCID_Id', 0x0002, projectlcid_id)
        projectlcid_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLCID_Size', 0x0004, projectlcid_size)
        # Lcid (4 bytes): An unsigned integer that specifies the LCID value for the VBA project. MUST be 0x00000409.
        self.lcid = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLCID_Lcid', 0x409, self.lcid)

        # PROJECTLCIDINVOKE Record
        # Specifies an LCID value used for Invoke calls on an Automation server as specified in [MS-OAUT] section 3.1.4.4.
        projectlcidinvoke_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTLCIDINVOKE_Id', 0x0014, projectlcidinvoke_id)
        projectlcidinvoke_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLCIDINVOKE_Size', 0x0004, projectlcidinvoke_size)
        # LcidInvoke (4 bytes): An unsigned integer that specifies the LCID value used for Invoke calls. MUST be 0x00000409.
        self.lcidinvoke = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLCIDINVOKE_LcidInvoke', 0x409, self.lcidinvoke)

        # PROJECTCODEPAGE Record
        # Specifies the VBA project's code page.
        projectcodepage_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTCODEPAGE_Id', 0x0003, projectcodepage_id)
        projectcodepage_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTCODEPAGE_Size', 0x0002, projectcodepage_size)
        self.codepage = struct.unpack("<H", dir_stream.read(2))[0]
        self.codepage_name = codepages.get_codepage_name(self.codepage)
        log.debug('Project Code Page: %r - %s' % (self.codepage, self.codepage_name))
        self.codec = codepages.codepage2codec(self.codepage)
        log.debug('Python codec corresponding to code page %d: %s' % (self.codepage, self.codec))


        # PROJECTNAME Record
        # Specifies a unique VBA identifier as the name of the VBA project.
        projectname_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTNAME_Id', 0x0004, projectname_id)
        sizeof_projectname = struct.unpack("<L", dir_stream.read(4))[0]
        log.debug('Project name size: %d bytes' % sizeof_projectname)
        if sizeof_projectname < 1 or sizeof_projectname > 128:
            # TODO: raise an actual error? What is MS Office's behaviour?
            log.error("PROJECTNAME_SizeOfProjectName value not in range [1-128]: {0}".format(sizeof_projectname))
        projectname_bytes = dir_stream.read(sizeof_projectname)
        self.projectname = self.decode_bytes(projectname_bytes)


        # PROJECTDOCSTRING Record
        # Specifies the description for the VBA project.
        projectdocstring_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTDOCSTRING_Id', 0x0005, projectdocstring_id)
        projectdocstring_sizeof_docstring = struct.unpack("<L", dir_stream.read(4))[0]
        if projectdocstring_sizeof_docstring > 2000:
            log.error(
                "PROJECTDOCSTRING_SizeOfDocString value not in range: {0}".format(projectdocstring_sizeof_docstring))
        # DocString (variable): An array of SizeOfDocString bytes that specifies the description for the VBA project.
        # MUST contain MBCS characters encoded using the code page specified in PROJECTCODEPAGE (section 2.3.4.2.1.4).
        # MUST NOT contain null characters.
        docstring_bytes = dir_stream.read(projectdocstring_sizeof_docstring)
        self.docstring = self.decode_bytes(docstring_bytes)
        projectdocstring_reserved = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTDOCSTRING_Reserved', 0x0040, projectdocstring_reserved)
        projectdocstring_sizeof_docstring_unicode = struct.unpack("<L", dir_stream.read(4))[0]
        if projectdocstring_sizeof_docstring_unicode % 2 != 0:
            log.error("PROJECTDOCSTRING_SizeOfDocStringUnicode is not even")
        # DocStringUnicode (variable): An array of SizeOfDocStringUnicode bytes that specifies the description for the
        # VBA project. MUST contain UTF-16 characters. MUST NOT contain null characters.
        # MUST contain the UTF-16 encoding of DocString.
        docstring_unicode_bytes = dir_stream.read(projectdocstring_sizeof_docstring_unicode)
        self.docstring_unicode = docstring_unicode_bytes.decode('utf16', errors='replace')

        # PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
        projecthelpfilepath_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTHELPFILEPATH_Id', 0x0006, projecthelpfilepath_id)
        projecthelpfilepath_sizeof_helpfile1 = struct.unpack("<L", dir_stream.read(4))[0]
        if projecthelpfilepath_sizeof_helpfile1 > 260:
            log.error(
                "PROJECTHELPFILEPATH_SizeOfHelpFile1 value not in range: {0}".format(projecthelpfilepath_sizeof_helpfile1))
        projecthelpfilepath_helpfile1 = dir_stream.read(projecthelpfilepath_sizeof_helpfile1)
        projecthelpfilepath_reserved = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTHELPFILEPATH_Reserved', 0x003D, projecthelpfilepath_reserved)
        projecthelpfilepath_sizeof_helpfile2 = struct.unpack("<L", dir_stream.read(4))[0]
        if projecthelpfilepath_sizeof_helpfile2 != projecthelpfilepath_sizeof_helpfile1:
            log.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 does not equal PROJECTHELPFILEPATH_SizeOfHelpFile2")
        projecthelpfilepath_helpfile2 = dir_stream.read(projecthelpfilepath_sizeof_helpfile2)
        if projecthelpfilepath_helpfile2 != projecthelpfilepath_helpfile1:
            log.error("PROJECTHELPFILEPATH_HelpFile1 does not equal PROJECTHELPFILEPATH_HelpFile2")

        # PROJECTHELPCONTEXT Record
        projecthelpcontext_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTHELPCONTEXT_Id', 0x0007, projecthelpcontext_id)
        projecthelpcontext_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTHELPCONTEXT_Size', 0x0004, projecthelpcontext_size)
        projecthelpcontext_helpcontext = struct.unpack("<L", dir_stream.read(4))[0]
        unused = projecthelpcontext_helpcontext

        # PROJECTLIBFLAGS Record
        projectlibflags_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTLIBFLAGS_Id', 0x0008, projectlibflags_id)
        projectlibflags_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLIBFLAGS_Size', 0x0004, projectlibflags_size)
        projectlibflags_projectlibflags = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTLIBFLAGS_ProjectLibFlags', 0x0000, projectlibflags_projectlibflags)

        # PROJECTVERSION Record
        projectversion_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTVERSION_Id', 0x0009, projectversion_id)
        projectversion_reserved = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTVERSION_Reserved', 0x0004, projectversion_reserved)
        projectversion_versionmajor = struct.unpack("<L", dir_stream.read(4))[0]
        projectversion_versionminor = struct.unpack("<H", dir_stream.read(2))[0]
        unused = projectversion_versionmajor
        unused = projectversion_versionminor

        # PROJECTCONSTANTS Record
        projectconstants_id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTCONSTANTS_Id', 0x000C, projectconstants_id)
        projectconstants_sizeof_constants = struct.unpack("<L", dir_stream.read(4))[0]
        if projectconstants_sizeof_constants > 1015:
            log.error(
                "PROJECTCONSTANTS_SizeOfConstants value not in range: {0}".format(projectconstants_sizeof_constants))
        projectconstants_constants = dir_stream.read(projectconstants_sizeof_constants)
        projectconstants_reserved = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTCONSTANTS_Reserved', 0x003C, projectconstants_reserved)
        projectconstants_sizeof_constants_unicode = struct.unpack("<L", dir_stream.read(4))[0]
        if projectconstants_sizeof_constants_unicode % 2 != 0:
            log.error("PROJECTCONSTANTS_SizeOfConstantsUnicode is not even")
        projectconstants_constants_unicode = dir_stream.read(projectconstants_sizeof_constants_unicode)
        unused = projectconstants_constants
        unused = projectconstants_constants_unicode

        # array of REFERENCE records
        # Specifies a reference to an Automation type library or VBA project.
        check = None
        while True:
            check = struct.unpack("<H", dir_stream.read(2))[0]
            log.debug("reference type = {0:04X}".format(check))
            if check == 0x000F:
                break

            if check == 0x0016:
                # REFERENCENAME
                # Specifies the name of a referenced VBA project or Automation type library.
                reference_id = check
                reference_sizeof_name = struct.unpack("<L", dir_stream.read(4))[0]
                reference_name = dir_stream.read(reference_sizeof_name)
                log.debug('REFERENCE name: %s' % unicode2str(self.decode_bytes(reference_name)))
                reference_reserved = struct.unpack("<H", dir_stream.read(2))[0]
                # According to [MS-OVBA] 2.3.4.2.2.2 REFERENCENAME Record:
                # "Reserved (2 bytes): MUST be 0x003E. MUST be ignored."
                # So let's ignore it, otherwise it crashes on some files (issue #132)
                # PR #135 by @c1fe:
                # contrary to the specification I think that the unicode name
                # is optional. if reference_reserved is not 0x003E I think it
                # is actually the start of another REFERENCE record
                # at least when projectsyskind_syskind == 0x02 (Macintosh)
                if reference_reserved == 0x003E:
                    #if reference_reserved not in (0x003E, 0x000D):
                    #    raise UnexpectedDataError(dir_path, 'REFERENCE_Reserved',
                    #                              0x0003E, reference_reserved)
                    reference_sizeof_name_unicode = struct.unpack("<L", dir_stream.read(4))[0]
                    reference_name_unicode = dir_stream.read(reference_sizeof_name_unicode)
                    unused = reference_id
                    unused = reference_name
                    unused = reference_name_unicode
                    continue
                else:
                    check = reference_reserved
                    log.debug("reference type = {0:04X}".format(check))

            if check == 0x0033:
                # REFERENCEORIGINAL (followed by REFERENCECONTROL)
                # Specifies the identifier of the Automation type library the containing REFERENCECONTROL's
                # (section 2.3.4.2.2.3) twiddled type library was generated from.
                referenceoriginal_id = check
                referenceoriginal_sizeof_libidoriginal = struct.unpack("<L", dir_stream.read(4))[0]
                referenceoriginal_libidoriginal = dir_stream.read(referenceoriginal_sizeof_libidoriginal)
                log.debug('REFERENCE original lib id: %s' % unicode2str(self.decode_bytes(referenceoriginal_libidoriginal)))
                unused = referenceoriginal_id
                unused = referenceoriginal_libidoriginal
                continue

            if check == 0x002F:
                # REFERENCECONTROL
                # Specifies a reference to a twiddled type library and its extended type library.
                referencecontrol_id = check
                referencecontrol_sizetwiddled = struct.unpack("<L", dir_stream.read(4))[0]  # ignore
                referencecontrol_sizeof_libidtwiddled = struct.unpack("<L", dir_stream.read(4))[0]
                referencecontrol_libidtwiddled = dir_stream.read(referencecontrol_sizeof_libidtwiddled)
                log.debug('REFERENCE control twiddled lib id: %s' % unicode2str(self.decode_bytes(referencecontrol_libidtwiddled)))
                referencecontrol_reserved1 = struct.unpack("<L", dir_stream.read(4))[0]  # ignore
                self.check_value('REFERENCECONTROL_Reserved1', 0x0000, referencecontrol_reserved1)
                referencecontrol_reserved2 = struct.unpack("<H", dir_stream.read(2))[0]  # ignore
                self.check_value('REFERENCECONTROL_Reserved2', 0x0000, referencecontrol_reserved2)
                unused = referencecontrol_id
                unused = referencecontrol_sizetwiddled
                unused = referencecontrol_libidtwiddled
                # optional field
                check2 = struct.unpack("<H", dir_stream.read(2))[0]
                if check2 == 0x0016:
                    referencecontrol_namerecordextended_id = check
                    referencecontrol_namerecordextended_sizeof_name = struct.unpack("<L", dir_stream.read(4))[0]
                    referencecontrol_namerecordextended_name = dir_stream.read(
                        referencecontrol_namerecordextended_sizeof_name)
                    log.debug('REFERENCE control name record extended: %s' % unicode2str(
                        self.decode_bytes(referencecontrol_namerecordextended_name)))
                    referencecontrol_namerecordextended_reserved = struct.unpack("<H", dir_stream.read(2))[0]
                    if referencecontrol_namerecordextended_reserved == 0x003E:
                        referencecontrol_namerecordextended_sizeof_name_unicode = struct.unpack("<L", dir_stream.read(4))[0]
                        referencecontrol_namerecordextended_name_unicode = dir_stream.read(
                            referencecontrol_namerecordextended_sizeof_name_unicode)
                        referencecontrol_reserved3 = struct.unpack("<H", dir_stream.read(2))[0]
                        unused = referencecontrol_namerecordextended_id
                        unused = referencecontrol_namerecordextended_name
                        unused = referencecontrol_namerecordextended_name_unicode
                    else:
                        referencecontrol_reserved3 = referencecontrol_namerecordextended_reserved
                else:
                    referencecontrol_reserved3 = check2

                self.check_value('REFERENCECONTROL_Reserved3', 0x0030, referencecontrol_reserved3)
                referencecontrol_sizeextended = struct.unpack("<L", dir_stream.read(4))[0]
                referencecontrol_sizeof_libidextended = struct.unpack("<L", dir_stream.read(4))[0]
                referencecontrol_libidextended = dir_stream.read(referencecontrol_sizeof_libidextended)
                referencecontrol_reserved4 = struct.unpack("<L", dir_stream.read(4))[0]
                referencecontrol_reserved5 = struct.unpack("<H", dir_stream.read(2))[0]
                referencecontrol_originaltypelib = dir_stream.read(16)
                referencecontrol_cookie = struct.unpack("<L", dir_stream.read(4))[0]
                unused = referencecontrol_sizeextended
                unused = referencecontrol_libidextended
                unused = referencecontrol_reserved4
                unused = referencecontrol_reserved5
                unused = referencecontrol_originaltypelib
                unused = referencecontrol_cookie
                continue

            if check == 0x000D:
                # REFERENCEREGISTERED
                # Specifies a reference to an Automation type library.
                referenceregistered_id = check
                referenceregistered_size = struct.unpack("<L", dir_stream.read(4))[0]
                referenceregistered_sizeof_libid = struct.unpack("<L", dir_stream.read(4))[0]
                referenceregistered_libid = dir_stream.read(referenceregistered_sizeof_libid)
                log.debug('REFERENCE registered lib id: %s' % unicode2str(self.decode_bytes(referenceregistered_libid)))
                referenceregistered_reserved1 = struct.unpack("<L", dir_stream.read(4))[0]
                self.check_value('REFERENCEREGISTERED_Reserved1', 0x0000, referenceregistered_reserved1)
                referenceregistered_reserved2 = struct.unpack("<H", dir_stream.read(2))[0]
                self.check_value('REFERENCEREGISTERED_Reserved2', 0x0000, referenceregistered_reserved2)
                unused = referenceregistered_id
                unused = referenceregistered_size
                unused = referenceregistered_libid
                continue

            if check == 0x000E:
                # REFERENCEPROJECT
                # Specifies a reference to an external VBA project.
                referenceproject_id = check
                referenceproject_size = struct.unpack("<L", dir_stream.read(4))[0]
                referenceproject_sizeof_libidabsolute = struct.unpack("<L", dir_stream.read(4))[0]
                referenceproject_libidabsolute = dir_stream.read(referenceproject_sizeof_libidabsolute)
                log.debug('REFERENCE project lib id absolute: %s' % unicode2str(self.decode_bytes(referenceproject_libidabsolute)))
                referenceproject_sizeof_libidrelative = struct.unpack("<L", dir_stream.read(4))[0]
                referenceproject_libidrelative = dir_stream.read(referenceproject_sizeof_libidrelative)
                log.debug('REFERENCE project lib id relative: %s' % unicode2str(self.decode_bytes(referenceproject_libidrelative)))
                referenceproject_majorversion = struct.unpack("<L", dir_stream.read(4))[0]
                referenceproject_minorversion = struct.unpack("<H", dir_stream.read(2))[0]
                unused = referenceproject_id
                unused = referenceproject_size
                unused = referenceproject_libidabsolute
                unused = referenceproject_libidrelative
                unused = referenceproject_majorversion
                unused = referenceproject_minorversion
                continue

            log.error('invalid or unknown check Id {0:04X}'.format(check))
            # raise an exception instead of stopping abruptly (issue #180)
            raise UnexpectedDataError(dir_path, 'reference type', (0x0F, 0x16, 0x33, 0x2F, 0x0D, 0x0E), check)
            #sys.exit(0)

    def check_value(self, name, expected, value):
        if expected != value:
            if self.relaxed:
                # It happens quite often that some values do not strictly follow
                # the MS-OVBA specifications, and this does not prevent the VBA
                # code from being extracted, so here we only raise a warning:
                log.warning("invalid value for {0} expected {1:04X} got {2:04X}"
                          .format(name, expected, value))
            else:
                raise UnexpectedDataError(self.dir_path, name, expected, value)

    def parse_project_stream(self):
        """
        Parse the PROJECT stream from the VBA project
        :return:
        """
        # Open the PROJECT stream:
        # reference: [MS-OVBA] 2.3.1 PROJECT Stream
        project_stream = self.ole.openstream(self.project_path)

        # sample content of the PROJECT stream:

        ##    ID="{5312AC8A-349D-4950-BDD0-49BE3C4DD0F0}"
        ##    Document=ThisDocument/&H00000000
        ##    Module=NewMacros
        ##    Name="Project"
        ##    HelpContextID="0"
        ##    VersionCompatible32="393222000"
        ##    CMG="F1F301E705E705E705E705"
        ##    DPB="8F8D7FE3831F2020202020"
        ##    GC="2D2FDD81E51EE61EE6E1"
        ##
        ##    [Host Extender Info]
        ##    &H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000
        ##    &H00000002={000209F2-0000-0000-C000-000000000046};Word8.0;&H00000000
        ##
        ##    [Workspace]
        ##    ThisDocument=22, 29, 339, 477, Z
        ##    NewMacros=-4, 42, 832, 510, C

        self.module_ext = {}

        for line in project_stream:
            line = self.decode_bytes(line)
            log.debug('PROJECT: %r' % line)
            line = line.strip()
            if '=' in line:
                # split line at the 1st equal sign:
                name, value = line.split('=', 1)
                # looking for code modules
                # add the code module as a key in the dictionary
                # the value will be the extension needed later
                # The value is converted to lowercase, to allow case-insensitive matching (issue #3)
                value = value.lower()
                if name == 'Document':
                    # split value at the 1st slash, keep 1st part:
                    value = value.split('/', 1)[0]
                    self.module_ext[value] = CLASS_EXTENSION
                elif name == 'Module':
                    self.module_ext[value] = MODULE_EXTENSION
                elif name == 'Class':
                    self.module_ext[value] = CLASS_EXTENSION
                elif name == 'BaseClass':
                    self.module_ext[value] = FORM_EXTENSION

    def parse_modules(self):
        dir_stream = self.dir_stream
        # projectmodules_id has already been read by the previous loop = 0x000F
        # projectmodules_id = check  #struct.unpack("<H", dir_stream.read(2))[0]
        # self.check_value('PROJECTMODULES_Id', 0x000F, projectmodules_id)
        projectmodules_size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTMODULES_Size', 0x0002, projectmodules_size)
        self.modules_count = struct.unpack("<H", dir_stream.read(2))[0]
        _id = struct.unpack("<H", dir_stream.read(2))[0]
        self.check_value('PROJECTMODULES_ProjectCookieRecord_Id', 0x0013, _id)
        size = struct.unpack("<L", dir_stream.read(4))[0]
        self.check_value('PROJECTMODULES_ProjectCookieRecord_Size', 0x0002, size)
        projectcookierecord_cookie = struct.unpack("<H", dir_stream.read(2))[0]
        unused = projectcookierecord_cookie

        log.debug("parsing {0} modules".format(self.modules_count))
        for module_index in xrange(0, self.modules_count):
            module = VBA_Module(self, self.dir_stream, module_index=module_index)
            self.modules.append(module)
            yield (module.code_path, module.filename_str, module.code_str)
        _ = unused   # make pylint happy: now variable "unused" is being used ;-)
        return

    def decode_bytes(self, bytes_string, errors='replace'):
        """
        Decode a bytes string to a unicode string, using the project code page
        :param bytes_string: bytes, bytes string to be decoded
        :param errors: str, mode to handle unicode conversion errors
        :return: str/unicode, decoded string
        """
        return bytes_string.decode(self.codec, errors=errors)



def _extract_vba(ole, vba_root, project_path, dir_path, relaxed=True):
    """
    Extract VBA macros from an OleFileIO object.
    Internal function, do not call directly.

    vba_root: path to the VBA root storage, containing the VBA storage and the PROJECT stream
    vba_project: path to the PROJECT stream
    :param relaxed: If True, only create info/debug log entry if data is not as expected
                    (e.g. opening substream fails); if False, raise an error in this case
    This is a generator, yielding (stream path, VBA filename, VBA source code) for each VBA code stream
    """
    log.debug('relaxed is %s' % relaxed)

    project = VBA_Project(ole, vba_root, project_path, dir_path, relaxed)
    project.parse_project_stream()

    for code_path, filename, code_data in project.parse_modules():
        yield (code_path, filename, code_data)


def vba_collapse_long_lines(vba_code):
    """
    Parse a VBA module code to detect continuation line characters (underscore) and
    collapse split lines. Continuation line characters are replaced by spaces.

    :param vba_code: str, VBA module code
    :return: str, VBA module code with long lines collapsed
    """
    # TODO: use a regex instead, to allow whitespaces after the underscore?
    try:
        vba_code = vba_code.replace(' _\r\n', ' ')
        vba_code = vba_code.replace(' _\r', ' ')
        vba_code = vba_code.replace(' _\n', ' ')
    except:
        log.exception('type(vba_code)=%s' % type(vba_code))
        raise
    return vba_code


def filter_vba(vba_code):
    """
    Filter VBA source code to remove the first lines starting with "Attribute VB_",
    which are automatically added by MS Office and not displayed in the VBA Editor.
    This should only be used when displaying source code for human analysis.

    Note: lines are not filtered if they contain a colon, because it could be
    used to hide malicious instructions.

    :param vba_code: str, VBA source code
    :return: str, filtered VBA source code
    """
    vba_lines = vba_code.splitlines()
    start = 0
    for line in vba_lines:
        if line.startswith("Attribute VB_") and not ':' in line:
            start += 1
        else:
            break
    #TODO: also remove empty lines?
    vba = '\n'.join(vba_lines[start:])
    return vba


def detect_autoexec(vba_code, obfuscation=None):
    """
    Detect if the VBA code contains keywords corresponding to macros running
    automatically when triggered by specific actions (e.g. when a document is
    opened or closed).

    :param vba_code: str, VBA source code
    :param obfuscation: None or str, name of obfuscation to be added to description
    :return: list of str tuples (keyword, description)
    """
    #TODO: merge code with detect_suspicious
    # case-insensitive search
    #vba_code = vba_code.lower()
    results = []
    obf_text = ''
    if obfuscation:
        obf_text = ' (obfuscation: %s)' % obfuscation
    # 1) simple strings, without regex
    for description, keywords in AUTOEXEC_KEYWORDS.items():
        for keyword in keywords:
            #TODO: if keyword is already a compiled regex, use it as-is
            # search using regex to detect word boundaries:
            match = re.search(r'(?i)\b' + re.escape(keyword) + r'\b', vba_code)
            if match:
                found_keyword = match.group()
                results.append((found_keyword, description + obf_text))
    # 2) regex
    for description, keywords in AUTOEXEC_KEYWORDS_REGEX.items():
        for keyword in keywords:
            #TODO: if keyword is already a compiled regex, use it as-is
            # search using regex to detect word boundaries:
            match = re.search(r'(?i)\b' + keyword + r'\b', vba_code)
            if match:
                found_keyword = match.group()
                results.append((found_keyword, description + obf_text))
    return results


def detect_suspicious(vba_code, obfuscation=None):
    """
    Detect if the VBA code contains suspicious keywords corresponding to
    potential malware behaviour.

    :param vba_code: str, VBA source code
    :param obfuscation: None or str, name of obfuscation to be added to description
    :return: list of str tuples (keyword, description)
    """
    # case-insensitive search
    #vba_code = vba_code.lower()
    results = []
    obf_text = ''
    if obfuscation:
        obf_text = ' (obfuscation: %s)' % obfuscation
    for description, keywords in SUSPICIOUS_KEYWORDS.items():
        for keyword in keywords:
            # search using regex to detect word boundaries:
            # note: each keyword must be escaped if it contains special chars such as '\'
            match = re.search(r'(?i)\b' + re.escape(keyword) + r'\b', vba_code)
            if match:
                found_keyword = match.group()
                results.append((found_keyword, description + obf_text))
    for description, keywords in SUSPICIOUS_KEYWORDS_REGEX.items():
        for keyword in keywords:
            # search using regex to detect word boundaries:
            # note: each keyword must NOT be escaped because it is an actual regex
            match = re.search(r'(?i)\b' + keyword + r'\b', vba_code)
            if match:
                found_keyword = match.group()
                results.append((found_keyword, description + obf_text))
    for description, keywords in SUSPICIOUS_KEYWORDS_NOREGEX.items():
        for keyword in keywords:
            if keyword.lower() in vba_code:
                # avoid reporting backspace chars out of plain VBA code:
                if not(keyword=='\b' and obfuscation is not None):
                    results.append((keyword, description + obf_text))
    return results


def detect_patterns(vba_code, obfuscation=None):
    """
    Detect if the VBA code contains specific patterns such as IP addresses,
    URLs, e-mail addresses, executable file names, etc.

    :param vba_code: str, VBA source code
    :return: list of str tuples (pattern type, value)
    """
    results = []
    found = set()
    obf_text = ''
    if obfuscation:
        obf_text = ' (obfuscation: %s)' % obfuscation
    for pattern_type, pattern_re in RE_PATTERNS:
        for match in pattern_re.finditer(vba_code):
            value = match.group()
            exclude_pattern_found = False
            for url_exclude_pattern in EXCLUDE_URLS_PATTERNS:
                if value.startswith(url_exclude_pattern):
                    exclude_pattern_found = True
            if value not in found and not exclude_pattern_found:
                results.append((pattern_type + obf_text, value))
                found.add(value)
    return results


def detect_hex_strings(vba_code):
    """
    Detect if the VBA code contains strings encoded in hexadecimal.

    :param vba_code: str, VBA source code
    :return: list of str tuples (encoded string, decoded string)
    """
    results = []
    found = set()
    for match in re_hex_string.finditer(vba_code):
        value = match.group()
        if value not in found:
            decoded = bytes2str(binascii.unhexlify(value))
            results.append((value, decoded))
            found.add(value)
    return results


def detect_base64_strings(vba_code):
    """
    Detect if the VBA code contains strings encoded in base64.

    :param vba_code: str, VBA source code
    :return: list of str tuples (encoded string, decoded string)
    """
    #TODO: avoid matching simple hex strings as base64?
    results = []
    found = set()
    for match in re_base64_string.finditer(vba_code):
        # extract the base64 string without quotes:
        value = match.group().strip('"')
        # check it is not just a hex string:
        if not re_nothex_check.search(value):
            continue
        # only keep new values and not in the whitelist:
        if value not in found and value.lower() not in BASE64_WHITELIST:
            try:
                decoded = bytes2str(base64.b64decode(value))
                results.append((value, decoded))
                found.add(value)
            except (TypeError, ValueError) as exc:
                log.debug('Failed to base64-decode (%s)' % exc)
                # if an exception occurs, it is likely not a base64-encoded string
    return results

# DridexUrlDecode written by James Habben
# Originally published on https://github.com/JamesHabben/MalwareStuff
# included here with James' permission
# 2015-01-27 Slight modifications from Philippe Lagadec (PL) to use it from olevba

def StripChars (input) :
    result = ''
    for c in input :
        if c.isdigit() :
            result += c
    return int(result)

def StripCharsWithZero (input) :
    result = ''
    for c in input :
        if c.isdigit() :
            result += c
        else:
            result += '0'
    return int(result)

def DridexUrlDecode (inputText) :
    work = inputText[4:-4]
    strKeyEnc = StripCharsWithZero(work[(len(work) / 2) - 2: (len(work) / 2)])
    strKeySize = StripCharsWithZero(work[(len(work) / 2): (len(work) / 2) + 2])
    nCharSize = strKeySize - strKeyEnc
    work = work[:(len(work) / 2) - 2] + work[(len(work) / 2) + 2:]
    strKeyEnc2 = StripChars(work[(len(work) / 2) - (nCharSize/2): (len(work) / 2) + (nCharSize/2)])
    work = work[:(len(work) / 2) - (nCharSize/2)] + work[(len(work) / 2) + (nCharSize/2):]
    work_split = [work[i:i+nCharSize] for i in range(0, len(work), nCharSize)]
    decoded = ''
    for group in work_split:
        # sys.stdout.write(chr(StripChars(group)/strKeyEnc2))
        decoded += chr(StripChars(group)/strKeyEnc2)
    return decoded

# DridexUrlDecode("C3iY1epSRGe6q8g15xStVesdG717MAlg2H4hmV1vkL6Glnf0cknj")
# DridexUrlDecode("HLIY3Nf3z2k8jD37h1n2OM3N712DGQ3c5M841RZ8C5e6P1C50C4ym1oF504WyV182p4mJ16cK9Z61l47h2dU1rVB5V681sFY728i16H3E2Qm1fn47y2cgAo156j8T1s600hukKO1568X1xE4Z7d2q17jvcwgk816Yz32o9Q216Mpr0B01vcwg856a17b9j2zAmWf1536B1t7d92rI1FZ5E36Pu1jl504Z34tm2R43i55Lg2F3eLE3T28lLX1D504348Goe8Gbdp37w443ADy36X0h14g7Wb2G3u584kEG332Ut8ws3wO584pzSTf")
# DridexUrlDecode("YNPH1W47E211z3P6142cM4115K2J1696CURf1712N1OCJwc0w6Z16840Z1r600W16Z3273k6SR16Bf161Q92a016Vr16V1pc")


def detect_dridex_strings(vba_code):
    """
    Detect if the VBA code contains strings obfuscated with a specific algorithm found in Dridex samples.

    :param vba_code: str, VBA source code
    :return: list of str tuples (encoded string, decoded string)
    """
    results = []
    found = set()
    for match in re_dridex_string.finditer(vba_code):
        value = match.group()[1:-1]
        # check it is not just a hex string:
        if not re_nothex_check.search(value):
            continue
        if value not in found:
            try:
                decoded = bytes2str(DridexUrlDecode(value))
                results.append((value, decoded))
                found.add(value)
            except Exception as exc:
                log.debug('Failed to Dridex-decode (%s)' % exc)
                # if an exception occurs, it is likely not a dridex-encoded string
    return results


def detect_vba_strings(vba_code):
    """
    Detect if the VBA code contains strings obfuscated with VBA expressions
    using keywords such as Chr, Asc, Val, StrReverse, etc.

    :param vba_code: str, VBA source code
    :return: list of str tuples (encoded string, decoded string)
    """
    # TODO: handle exceptions
    results = []
    found = set()
    # IMPORTANT: to extract the actual VBA expressions found in the code,
    #            we must expand tabs to have the same string as pyparsing.
    #            Otherwise, start and end offsets are incorrect.
    vba_code = vba_code.expandtabs()
    # Split the VBA code line by line to avoid MemoryError on large scripts:
    for vba_line in vba_code.splitlines():
        for tokens, start, end in vba_expr_str.scanString(vba_line):
            encoded = vba_line[start:end]
            decoded = tokens[0]
            if isinstance(decoded, VbaExpressionString):
                # This is a VBA expression, not a simple string
                # print 'VBA EXPRESSION: encoded=%r => decoded=%r' % (encoded, decoded)
                # remove parentheses and quotes from original string:
                # if encoded.startswith('(') and encoded.endswith(')'):
                #     encoded = encoded[1:-1]
                # if encoded.startswith('"') and encoded.endswith('"'):
                #     encoded = encoded[1:-1]
                # avoid duplicates and simple strings:
                if encoded not in found and decoded != encoded:
                    results.append((encoded, decoded))
                    found.add(encoded)
            # else:
                # print 'VBA STRING: encoded=%r => decoded=%r' % (encoded, decoded)
    return results


def json2ascii(json_obj, encoding='utf8', errors='replace'):
    """
    ensure there is no unicode in json and all strings are safe to decode

    works recursively, decodes and re-encodes every string to/from unicode
    to ensure there will be no trouble in loading the dumped json output
    """
    if json_obj is None:
        pass
    elif isinstance(json_obj, (bool, int, float)):
        pass
    elif isinstance(json_obj, str):
        if PYTHON2:
            # de-code and re-encode
            dencoded = json_obj.decode(encoding, errors).encode(encoding, errors)
            if dencoded != json_obj:
                log.debug('json2ascii: replaced: {0} (len {1})'
                         .format(json_obj, len(json_obj)))
                log.debug('json2ascii:     with: {0} (len {1})'
                         .format(dencoded, len(dencoded)))
            return dencoded
        else:
            # on Python 3, just keep Unicode strings as-is:
            return json_obj
    # pylint: disable-next=possibly-used-before-assignment
    elif isinstance(json_obj, unicode) and PYTHON2:
        # On Python 2, encode unicode to bytes:
        json_obj_bytes = json_obj.encode(encoding, errors)
        log.debug('json2ascii: encode unicode: {0}'.format(json_obj_bytes))
        # cannot put original into logger
        # print 'original: ' json_obj
        return json_obj_bytes
    elif isinstance(json_obj, bytes) and not PYTHON2:
        # On Python 3, decode bytes to unicode str
        json_obj_str = json_obj.decode(encoding, errors)
        log.debug('json2ascii: encode unicode: {0}'.format(json_obj_str))
        # cannot put original into logger
        # print 'original: ' json_obj
        return json_obj_str
    elif isinstance(json_obj, dict):
        for key in json_obj:
            json_obj[key] = json2ascii(json_obj[key])
    elif isinstance(json_obj, (list,tuple)):
        for item in json_obj:
            item = json2ascii(item)
    else:
        log.debug('unexpected type in json2ascii: {0} -- leave as is'
                  .format(type(json_obj)))
    return json_obj


def print_json(json_dict=None, _json_is_first=False, **json_parts):
    """ line-wise print of json.dumps(json2ascii(..)) with options and indent+1

    can use in two ways:
    (1) print_json(some_dict)
    (2) print_json(key1=value1, key2=value2, ...)

    This is compatible with :py:mod:`oletools.common.log_helper`: log messages
    can be mixed if arg `use_json` was `True` in
    :py:func:`log_helper.enable_logging` provided this function is called
    before the first "regular" logging with `_json_is_first=True` (and
    non-empty input) but after log_helper.enable_logging.
    """
    if json_dict and json_parts:
        raise ValueError('Invalid json argument: want either single dict or '
                         'key=value parts but got both)')
    elif (json_dict is not None) and (not isinstance(json_dict, dict)):
        raise ValueError('Invalid json argument: want either single dict or '
                         'key=value parts but got {0} instead of dict)'
                         .format(type(json_dict)))
    if json_parts:
        json_dict = json_parts

    lines = json.dumps(json2ascii(json_dict), check_circular=False,
                       indent=4, ensure_ascii=False).splitlines()
    if not lines:
        return

    if _json_is_first:
        print('      ' + lines[0])
    else:
        print(',     ' + lines[0])

    for line in lines[1:]:
        print('      ' + line.rstrip())


class VBA_Scanner(object):
    """
    Class to scan the source code of a VBA module to find obfuscated strings,
    suspicious keywords, IOCs, auto-executable macros, etc.
    """

    def __init__(self, vba_code):
        """
        VBA_Scanner constructor

        :param vba_code: str, VBA source code to be analyzed
        """
        # join long lines ending with " _":
        self.code = vba_collapse_long_lines(vba_code)
        self.code_hex = ''
        self.code_hex_rev = ''
        self.code_rev_hex = ''
        self.code_base64 = ''
        self.code_dridex = ''
        self.code_vba = ''
        self.strReverse = None
        # results = None before scanning, then a list of tuples after scanning
        self.results = None
        self.autoexec_keywords = None
        self.suspicious_keywords = None
        self.iocs = None
        self.hex_strings = None
        self.base64_strings = None
        self.dridex_strings = None
        self.vba_strings = None


    def scan(self, include_decoded_strings=False, deobfuscate=False):
        """
        Analyze the provided VBA code to detect suspicious keywords,
        auto-executable macros, IOC patterns, obfuscation patterns
        such as hex-encoded strings.

        :param include_decoded_strings: bool, if True, all encoded strings will be included with their decoded content.
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
        :return: list of tuples (type, keyword, description)
        (type = 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String' or 'Dridex String')
        """
        # First, detect and extract hex-encoded strings:
        self.hex_strings = detect_hex_strings(self.code)
        # detect if the code contains StrReverse:
        self.strReverse = False
        if 'strreverse' in self.code.lower(): self.strReverse = True
        # Then append the decoded strings to the VBA code, to detect obfuscated IOCs and keywords:
        for encoded, decoded in self.hex_strings:
            self.code_hex += '\n' + decoded
            # if the code contains "StrReverse", also append the hex strings in reverse order:
            if self.strReverse:
                # StrReverse after hex decoding:
                self.code_hex_rev += '\n' + decoded[::-1]
                # StrReverse before hex decoding:
                self.code_rev_hex += '\n' + bytes2str(binascii.unhexlify(encoded[::-1]))
                #example: https://malwr.com/analysis/NmFlMGI4YTY1YzYyNDkwNTg1ZTBiZmY5OGI3YjlhYzU/
        #TODO: also append the full code reversed if StrReverse? (risk of false positives?)
        # Detect Base64-encoded strings
        self.base64_strings = detect_base64_strings(self.code)
        for encoded, decoded in self.base64_strings:
            self.code_base64 += '\n' + decoded
        # Detect Dridex-encoded strings
        self.dridex_strings = detect_dridex_strings(self.code)
        for encoded, decoded in self.dridex_strings:
            self.code_dridex += '\n' + decoded
        # Detect obfuscated strings in VBA expressions
        if deobfuscate:
            self.vba_strings = detect_vba_strings(self.code)
        else:
            self.vba_strings = []
        for encoded, decoded in self.vba_strings:
            self.code_vba += '\n' + decoded
        results = []
        self.autoexec_keywords = []
        self.suspicious_keywords = []
        self.iocs = []

        for code, obfuscation in (
                (self.code, None),
                (self.code_hex, 'Hex'),
                (self.code_hex_rev, 'Hex+StrReverse'),
                (self.code_rev_hex, 'StrReverse+Hex'),
                (self.code_base64, 'Base64'),
                (self.code_dridex, 'Dridex'),
                (self.code_vba, 'VBA expression'),
        ):
            self.autoexec_keywords += detect_autoexec(code, obfuscation)
            self.suspicious_keywords += detect_suspicious(code, obfuscation)
            self.iocs += detect_patterns(code, obfuscation)

        # If hex-encoded strings were discovered, add an item to suspicious keywords:
        if self.hex_strings:
            self.suspicious_keywords.append(('Hex Strings',
                                             'Hex-encoded strings were detected, may be used to obfuscate strings (option --decode to see all)'))
        if self.base64_strings:
            self.suspicious_keywords.append(('Base64 Strings',
                                             'Base64-encoded strings were detected, may be used to obfuscate strings (option --decode to see all)'))
        if self.dridex_strings:
            self.suspicious_keywords.append(('Dridex Strings',
                                             'Dridex-encoded strings were detected, may be used to obfuscate strings (option --decode to see all)'))
        if self.vba_strings:
            self.suspicious_keywords.append(('VBA obfuscated Strings',
                                             'VBA string expressions were detected, may be used to obfuscate strings (option --decode to see all)'))
        # use a set to avoid duplicate keywords
        keyword_set = set()
        for keyword, description in self.autoexec_keywords:
            if keyword not in keyword_set:
                results.append(('AutoExec', keyword, description))
                keyword_set.add(keyword)
        keyword_set = set()
        for keyword, description in self.suspicious_keywords:
            if keyword not in keyword_set:
                results.append(('Suspicious', keyword, description))
                keyword_set.add(keyword)
        keyword_set = set()
        for pattern_type, value in self.iocs:
            if value not in keyword_set:
                results.append(('IOC', value, pattern_type))
                keyword_set.add(value)

        # include decoded strings only if they are printable or if --decode option:
        for encoded, decoded in self.hex_strings:
            if include_decoded_strings or is_printable(decoded):
                results.append(('Hex String', decoded, encoded))
        for encoded, decoded in self.base64_strings:
            if include_decoded_strings or is_printable(decoded):
                results.append(('Base64 String', decoded, encoded))
        for encoded, decoded in self.dridex_strings:
            if include_decoded_strings or is_printable(decoded):
                results.append(('Dridex string', decoded, encoded))
        for encoded, decoded in self.vba_strings:
            if include_decoded_strings or is_printable(decoded):
                results.append(('VBA string', decoded, encoded))
        self.results = results
        return results

    def scan_summary(self):
        """
        Analyze the provided VBA code to detect suspicious keywords,
        auto-executable macros, IOC patterns, obfuscation patterns
        such as hex-encoded strings.

        :return: tuple with the number of items found for each category:
            (autoexec, suspicious, IOCs, hex, base64, dridex, vba)
        """
        # avoid scanning the same code twice:
        if self.results is None:
            self.scan()
        return (len(self.autoexec_keywords), len(self.suspicious_keywords),
                len(self.iocs), len(self.hex_strings), len(self.base64_strings),
                len(self.dridex_strings), len(self.vba_strings))


def scan_vba(vba_code, include_decoded_strings, deobfuscate=False):
    """
    Analyze the provided VBA code to detect suspicious keywords,
    auto-executable macros, IOC patterns, obfuscation patterns
    such as hex-encoded strings.
    (shortcut for VBA_Scanner(vba_code).scan())

    :param vba_code: str, VBA source code to be analyzed
    :param include_decoded_strings: bool, if True all encoded strings will be included with their decoded content.
    :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
    :return: list of tuples (type, keyword, description)
        with type = 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String' or 'Dridex String'
    """
    return VBA_Scanner(vba_code).scan(include_decoded_strings, deobfuscate)


#=== CLASSES =================================================================

class VBA_Parser(object):
    """
    Class to parse MS Office files, to detect VBA macros and extract VBA source code
    """
    # TODO: relaxed is enabled by default temporarily, until a solution is found for issue #593

    def __init__(self, filename, data=None, container=None, relaxed=True, encoding=DEFAULT_API_ENCODING,
                 disable_pcode=False):
        """
        Constructor for VBA_Parser

        :param str filename: filename or path of file to parse, or file-like object

        :param bytes data: None or bytes str, if None the file will be read from disk (or from the file-like object).
            If data is provided as a bytes string, it will be parsed as the content of the file in memory,
            and not read from disk. Note: files must be read in binary mode, i.e. open(f, 'rb').

        :param str container: str, path and filename of container if the file is within
            a zip archive, None otherwise.

        :param bool relaxed: if True, treat mal-formed documents and missing streams more like MS office:
            do nothing; if False (default), raise errors in these cases

        :param str encoding: encoding for VBA source code and strings.
            Default: UTF-8 bytes strings on Python 2, unicode strings on Python 3 (None)

        raises a FileOpenError if all attempts to interpret the data header failed.
        """
        # TODO: filename should only be a string, data should be used for the file-like object
        # TODO: filename should be mandatory, optional data is a string or file-like object
        # TODO: also support olefile and zipfile as input
        if data is None:
            # open file from disk:
            _file = filename
            self.file_on_disk = True
        else:
            # file already read in memory, make it a file-like object for zipfile:
            _file = BytesIO(data)
            self.file_on_disk = False
        #self.file = _file
        self.ole_file = None
        self.ole_subfiles = []
        self.filename = filename
        self.container = container
        self.relaxed = relaxed
        self.type = None
        self.vba_projects = None
        self.vba_forms = None
        self.contains_vba_macros = None # will be set to True or False by detect_vba_macros
        self.contains_xlm_macros = None # will be set to True or False by detect_xlm_macros
        self.vba_code_all_modules = None # to store the source code of all modules
        # list of tuples for each module: (subfilename, stream_path, vba_filename, vba_code)
        self.modules = None
        # Analysis results: list of tuples (type, keyword, description) - See VBA_Scanner
        self.analysis_results = None
        # statistics for the scan summary and flags
        self.nb_macros = 0
        self.nb_autoexec = 0
        self.nb_suspicious = 0
        self.nb_iocs = 0
        self.nb_hexstrings = 0
        self.nb_base64strings = 0
        self.nb_dridexstrings = 0
        self.nb_vbastrings = 0
        #: Encoding for VBA source code and strings returned by all methods
        self.encoding = encoding
        self.xlm_macros = []
        self.no_xlm = False
        #: Output from pcodedmp, disassembly of the VBA P-code
        self.disable_pcode = disable_pcode
        self.pcodedmp_output = None
        #: Flag set to True/False if VBA stomping detected
        self.vba_stomping_detected = None
        # will be set to True or False by detect_is_encrypted method
        self.is_encrypted = False
        # TODO: those are disabled for now:
        self.xlm_macrosheet_found = False
        self.template_injection_found = False

        # call ftguess to identify file type:
        self.ftg = ftguess.FileTypeGuesser(self.filename, data=data)
        log.debug('ftguess: file type=%s - container=%s' % (self.ftg.ftype.name, self.ftg.container))
        # if filename is None:
        #     if isinstance(_file, basestring):
        #         if len(_file) < olefile.MINIMAL_OLEFILE_SIZE:
        #             self.filename = _file
        #         else:
        #             self.filename = '<file in bytes string>'
        #     else:
        #         self.filename = '<file-like object>'
        if olefile.isOleFile(_file):
            # This looks like an OLE file
            self.open_ole(_file)

            # if this worked, try whether it is a ppt file (special ole file)
            # TODO: instead of this we should have a function to test if it is a PPT (e.g. using ftguess)
            self.open_ppt()
        if self.type is None and zipfile.is_zipfile(_file):
            # Zip file, which may be an OpenXML document
            self.open_openxml(_file)
        if self.type is None:
            # read file from disk, check if it is a Word 2003 XML file (WordProcessingML), Excel 2003 XML,
            # or a plain text file containing VBA code
            if data is None:
                with open(filename, 'rb') as file_handle:
                    data = file_handle.read()
            # check if it is a Word 2003 XML file (WordProcessingML): must contain the namespace
            if b'http://schemas.microsoft.com/office/word/2003/wordml' in data:
                self.open_word2003xml(data)
            # check if it is a Word/PowerPoint 2007+ XML file (Flat OPC): must contain the namespace
            if b'http://schemas.microsoft.com/office/2006/xmlPackage' in data:
                self.open_flatopc(data)
            # store a lowercase version for the next tests:
            data_lowercase = data.lower()
            # check if it is a MHT file (MIME HTML, Word or Excel saved as "Single File Web Page"):
            # According to my tests, these files usually start with "MIME-Version: 1.0" on the 1st line
            # BUT Word accepts a blank line or other MIME headers inserted before,
            # and even whitespaces in between "MIME", "-", "Version" and ":". The version number is ignored.
            # And the line is case insensitive.
            # so we'll just check the presence of mime, version and multipart anywhere:
            if (self.type is None and
                b'mime' in data_lowercase and
                b'version' in data_lowercase and
                b'multipart' in data_lowercase and
                abs(data_lowercase.index(b'version') - data_lowercase.index(b'mime')) < 20):
                self.open_mht(data)
        #TODO: handle exceptions
        #TODO: Excel 2003 XML
            # Check whether this is rtf
            if rtfobj.is_rtf(data, treat_str_as_data=True):
                # Ignore RTF since it contains no macros and methods in here will not find macros
                # in embedded objects. run rtfobj and repeat on its output.
                msg = '%s is RTF, which cannot contain VBA Macros. Please use rtfobj to analyse it.' % self.filename
                log.info(msg)
                raise FileOpenError(msg)
            # Check if it is a SLK/SYLK file - https://en.wikipedia.org/wiki/SYmbolic_LinK_(SYLK)
            # It must start with "ID" in uppercase, no whitespace or newline allowed before by Excel:
            if data.startswith(b'ID'):
                self.open_slk(data)
            # Check if this is a plain text VBA or VBScript file:
            # To avoid scanning binary files, we simply check for some control chars:
            if self.type is None and b'\x00' not in data:
                self.open_text(data)
        if self.type is None:
            # At this stage, could not match a known format:
            msg = '%s is not a supported file type, cannot extract VBA Macros.' % self.filename
            log.info(msg)
            raise FileOpenError(msg)

    def open_ole(self, _file):
        """
        Open an OLE file
        :param _file: filename or file contents in a file object
        :return: nothing
        """
        log.info('Opening OLE file %s' % self.filename)
        try:
            # Open and parse the OLE file, using unicode for path names:
            self.ole_file = olefile.OleFileIO(_file, path_encoding=None)
            # set type only if parsing succeeds
            self.type = TYPE_OLE
        except (IOError, TypeError, ValueError) as exc:
            # TODO: handle OLE parsing exceptions
            log.info('Failed OLE parsing for file %r (%s)' % (self.filename, exc))
            log.debug('Trace:', exc_info=True)

    def open_openxml(self, _file):
        """
        Open an OpenXML file
        :param _file: filename or file contents in a file object
        :return: nothing
        """
        # This looks like a zip file, need to look for vbaProject.bin inside
        # It can be any OLE file inside the archive
        #...because vbaProject.bin can be renamed:
        # see http://www.decalage.info/files/JCV07_Lagadec_OpenDocument_OpenXML_v4_decalage.pdf#page=18
        log.info('Opening ZIP/OpenXML file %s' % self.filename)
        try:
            z = zipfile.ZipFile(_file)
            #TODO: check if this is actually an OpenXML file
            #TODO: if the zip file is encrypted, suggest to use the -z option, or try '-z infected' automatically
            # check each file within the zip if it is an OLE file, by reading its magic:
            for subfile in z.namelist():
                log.debug("OpenXML subfile {}".format(subfile))
                with z.open(subfile) as file_handle:
                    # found_ole = False
                    # template_injection_detected = False
                    # xml_macrosheet_found = False
                    magic = file_handle.read(len(olefile.MAGIC))
                    if magic == olefile.MAGIC:
                #         found_ole = True
                #     # in case we did not find an OLE file,
                #     # there could be a XLM macrosheet or a template injection attempt
                #     if not found_ole:
                #         read_all_file = file_handle.read()
                #         # try to detect template injection attempt
                #         # https://ired.team/offensive-security/initial-access/phishing-with-ms-office/inject-macros-from-a-remote-dotm-template-docx-with-macros
                #         subfile_that_can_contain_templates = "word/_rels/settings.xml.rels"
                #         if subfile == subfile_that_can_contain_templates:
                #             regex_template = b"Type=\"http://schemas\.openxmlformats\.org/officeDocument/\d{4}/relationships/attachedTemplate\"\s+Target=\"(.+?)\""
                #             template_injection_found = re.search(regex_template, read_all_file)
                #             if template_injection_found:
                #                 injected_template_url = template_injection_found.group(1).decode()
                #                 message = "Found injected template in subfile {}. Template URL: {}"\
                #                           "".format(subfile_that_can_contain_templates, injected_template_url)
                #                 log.info(message)
                #                 template_injection_detected = True
                #                 self.template_injection_found = True
                #         # try to find a XML macrosheet
                #         macro_sheet_footer = b"</xm:macrosheet>"
                #         len_macro_sheet_footer = len(macro_sheet_footer)
                #         last_bytes_to_check = read_all_file[-len_macro_sheet_footer:]
                #         if last_bytes_to_check == macro_sheet_footer:
                #             message = "Found XLM Macro in subfile: {}".format(subfile)
                #             log.info(message)
                #             xml_macrosheet_found = True
                #             self.xlm_macrosheet_found = True
                #
                # if found_ole or xml_macrosheet_found or template_injection_detected:
                        log.debug('Opening OLE file %s within zip' % subfile)
                        with z.open(subfile) as file_handle:
                            ole_data = file_handle.read()
                        try:
                            self.append_subfile(filename=subfile, data=ole_data)
                        except OlevbaBaseException as exc:
                            if self.relaxed:
                                log.info('%s is not a valid OLE file (%s)' % (subfile, exc))
                                log.debug('Trace:', exc_info=True)
                                continue
                            else:
                                raise SubstreamOpenError(self.filename, subfile,
                                                         exc)
            z.close()
            # set type only if parsing succeeds
            self.type = TYPE_OpenXML
        except OlevbaBaseException as exc:
            if self.relaxed:
                log.info('Error {0} caught in Zip/OpenXML parsing for file {1}'
                         .format(exc, self.filename))
                log.debug('Trace:', exc_info=True)
            else:
                raise
        except (RuntimeError, zipfile.BadZipfile, zipfile.LargeZipFile, IOError) as exc:
            # TODO: handle parsing exceptions
            log.info('Failed Zip/OpenXML parsing for file %r (%s)'
                          % (self.filename, exc))
            log.debug('Trace:', exc_info=True)

    def open_word2003xml(self, data):
        """
        Open a Word 2003 XML file
        :param data: file contents in a string or bytes
        :return: nothing
        """
        log.info('Opening Word 2003 XML file %s' % self.filename)
        try:
            # parse the XML content
            # TODO: handle XML parsing exceptions
            et = ET.fromstring(data)
            # find all the binData elements:
            for bindata in et.iter(TAG_BINDATA):
                # the binData content is an OLE container for the VBA project, compressed
                # using the ActiveMime/MSO format (zlib-compressed), and Base64 encoded.
                # get the filename:
                fname = bindata.get(ATTR_NAME, 'noname.mso')
                # decode the base64 activemime
                mso_data = binascii.a2b_base64(bindata.text)
                if is_mso_file(mso_data):
                    # decompress the zlib data stored in the MSO file, which is the OLE container:
                    # TODO: handle different offsets => separate function
                    try:
                        ole_data = mso_file_extract(mso_data)
                        self.append_subfile(filename=fname, data=ole_data)
                    except OlevbaBaseException as exc:
                        if self.relaxed:
                            log.info('Error parsing subfile {0}: {1}'
                                     .format(fname, exc))
                            log.debug('Trace:', exc_info=True)
                        else:
                            raise SubstreamOpenError(self.filename, fname, exc)
                else:
                    log.info('%s is not a valid MSO file' % fname)
            # set type only if parsing succeeds
            self.type = TYPE_Word2003_XML
        except OlevbaBaseException as exc:
            if self.relaxed:
                log.info('Failed XML parsing for file %r (%s)' % (self.filename, exc))
                log.debug('Trace:', exc_info=True)
            else:
                raise
        except Exception as exc:
            # TODO: differentiate exceptions for each parsing stage
            # (but ET is different libs, no good exception description in API)
            # found: XMLSyntaxError
            log.info('Failed XML parsing for file %r (%s)' % (self.filename, exc))
            log.debug('Trace:', exc_info=True)

    def open_flatopc(self, data):
        """
        Open a Word or PowerPoint 2007+ XML file, aka "Flat OPC"
        :param data: file contents in a string or bytes
        :return: nothing
        """
        log.info('Opening Flat OPC Word/PowerPoint XML file %s' % self.filename)
        try:
            # parse the XML content
            # TODO: handle XML parsing exceptions
            et = ET.fromstring(data)
            # TODO: check root node namespace and tag
            # find all the pkg:part elements:
            for pkgpart in et.iter(TAG_PKGPART):
                fname = pkgpart.get(ATTR_PKG_NAME, 'unknown')
                content_type = pkgpart.get(ATTR_PKG_CONTENTTYPE, 'unknown')
                if content_type == CTYPE_VBAPROJECT:
                    for bindata in pkgpart.iterfind(TAG_PKGBINDATA):
                        try:
                            ole_data = binascii.a2b_base64(bindata.text)
                            self.append_subfile(filename=fname, data=ole_data)
                        except OlevbaBaseException as exc:
                            if self.relaxed:
                                log.info('Error parsing subfile {0}: {1}'
                                         .format(fname, exc))
                                log.debug('Trace:', exc_info=True)
                            else:
                                raise SubstreamOpenError(self.filename, fname, exc)
            # set type only if parsing succeeds
            self.type = TYPE_FlatOPC_XML
        except OlevbaBaseException as exc:
            if self.relaxed:
                log.info('Failed XML parsing for file %r (%s)' % (self.filename, exc))
                log.debug('Trace:', exc_info=True)
            else:
                raise
        except Exception as exc:
            # TODO: differentiate exceptions for each parsing stage
            # (but ET is different libs, no good exception description in API)
            # found: XMLSyntaxError
            log.info('Failed XML parsing for file %r (%s)' % (self.filename, exc))
            log.debug('Trace:', exc_info=True)

    def open_mht(self, data):
        """
        Open a MHTML file
        :param data: file contents in a string or bytes
        :return: nothing
        """
        log.info('Opening MHTML file %s' % self.filename)
        try:
            # parse the MIME content
            # remove any leading whitespace or newline (workaround for issue in email package)
            stripped_data = data.lstrip(b'\r\n\t ')
            # strip any junk from the beginning of the file
            # (issue #31 fix by Greg C - gdigreg)
            # TODO: improve keywords to avoid false positives
            mime_offset = stripped_data.find(b'MIME')
            content_offset = stripped_data.find(b'Content')
            # if "MIME" is found, and located before "Content":
            if -1 < mime_offset <= content_offset:
                stripped_data = stripped_data[mime_offset:]
            # else if "Content" is found, and before "MIME"
            # TODO: can it work without "MIME" at all?
            elif content_offset > -1:
                stripped_data = stripped_data[content_offset:]
            # TODO: quick and dirty fix: insert a standard line with MIME-Version header?
            # monkeypatch email to fix issue #32:
            # allow header lines without ":"
            oldHeaderRE = email.feedparser.headerRE
            loosyHeaderRE = re.compile(r'^(From |[\041-\071\073-\176]{1,}:?|[\t ])')
            email.feedparser.headerRE = loosyHeaderRE
            try:
                if PYTHON2:
                    mhtml = email.message_from_string(stripped_data)
                else:
                    # on Python 3, need to use message_from_bytes instead:
                    mhtml = email.message_from_bytes(stripped_data)
            finally:
                email.feedparser.headerRE = oldHeaderRE
            # find all the attached files:
            for part in mhtml.walk():
                content_type = part.get_content_type()  # always returns a value
                fname = part.get_filename(None)  # returns None if it fails
                # TODO: get content-location if no filename
                log.debug('MHTML part: filename=%r, content-type=%r' % (fname, content_type))
                part_data = part.get_payload(decode=True)
                # VBA macros are stored in a binary file named "editdata.mso".
                # the data content is an OLE container for the VBA project, compressed
                # using the ActiveMime/MSO format (zlib-compressed), and Base64 encoded.
                # decompress the zlib data starting at offset 0x32, which is the OLE container:
                # check ActiveMime header:
                if isinstance(part_data, bytes) and is_mso_file(part_data):
                    log.debug('Found ActiveMime header, decompressing MSO container')
                    try:
                        ole_data = mso_file_extract(part_data)

                        # TODO: check if it is actually an OLE file
                        # TODO: get the MSO filename from content_location?
                        self.append_subfile(filename=fname, data=ole_data)
                    except OlevbaBaseException as exc:
                        if self.relaxed:
                            log.info('%s does not contain a valid OLE file (%s)'
                                      % (fname, exc))
                            log.debug('Trace:', exc_info=True)
                            # TODO: bug here - need to split in smaller functions/classes?
                        else:
                            raise SubstreamOpenError(self.filename, fname, exc)
                else:
                    log.debug('type(part_data) = %s' % type(part_data))
                    try:
                        log.debug('part_data[0:20] = %r' % part_data[0:20])
                    except TypeError as err:
                        log.debug('part_data has no __getitem__')
            # set type only if parsing succeeds
            self.type = TYPE_MHTML
        except OlevbaBaseException:
            raise
        except Exception:
            log.info('Failed MIME parsing for file %r - %s'
                          % (self.filename, MSG_OLEVBA_ISSUES))
            log.debug('Trace:', exc_info=True)

    def open_ppt(self):
        """ try to interpret self.ole_file as PowerPoint 97-2003 using PptParser

        Although self.ole_file is a valid olefile.OleFileIO, we set
        self.ole_file = None in here and instead set self.ole_subfiles to the
        VBA ole streams found within the main ole file. That makes most of the
        code below treat this like an OpenXML file and only look at the
        ole_subfiles (except find_vba_* which needs to explicitly check for
        self.type)
        """

        log.info('Check whether OLE file is PPT')
        try:
            ppt = ppt_parser.PptParser(self.ole_file, fast_fail=True)
            for vba_data in ppt.iter_vba_data():    # pylint: disable=no-value-for-parameter
                self.append_subfile(None, vba_data, container='PptParser')
            log.info('File is PPT')
            self.ole_file.close()  # just in case
            self.ole_file = None   # required to make other methods look at ole_subfiles
            self.type = TYPE_PPT
        except (ppt_parser.PptUnexpectedData, ValueError) as exc:
            if self.container == 'PptParser':
                # this is a subfile of a ppt --> to be expected that is no ppt
                log.debug('PPT subfile is not a PPT file')
            else:
                log.debug("File appears not to be a ppt file (%s)" % exc)


    def open_slk(self, data):
        """
        Open a SLK file, which may contain XLM/Excel 4 macros
        :param data: file contents in a bytes string
        :return: nothing
        """
        # TODO: Those results should be stored as XLM macros, not VBA
        log.info('Opening SLK file %s' % self.filename)
        xlm_macro_found = False
        xlm_macros = []
        xlm_macros.append('Formulas and XLM/Excel 4 macros extracted from SLK file:')
        for line in data.splitlines(False):
            if line.startswith(b'O'):
                # Option: "O;E" indicates a macro sheet, must appear before NN and C rows
                for s in line.split(b';'):
                    if s.startswith(b'E'):
                        xlm_macro_found = True
                        log.debug('SLK parser: found macro sheet')
            elif line.startswith(b'NN') and xlm_macro_found:
                # Name that can trigger a macro, for example "Auto_Open"
                for s in line.split(b';'):
                    if s.startswith(b'N') and s.strip() != b'NN':
                        xlm_macros.append('Named cell: %s' % bytes2str(s[1:]))
            elif line.startswith(b'C') and xlm_macro_found:
                # Cell
                for s in line.split(b';'):
                    if s.startswith(b'E'):
                        xlm_macros.append('Formula or Macro: %s' % bytes2str(s[1:]))
        if xlm_macro_found:
            self.contains_xlm_macros = True
            self.xlm_macros = xlm_macros
        self.type = TYPE_SLK


    def open_text(self, data):
        """
        Open a text file containing VBA or VBScript source code
        :param data: file contents in a string or bytes
        :return: nothing
        """
        log.info('Opening text file %s' % self.filename)
        # directly store the source code:
        # On Python 2, store it as a raw bytes string
        # On Python 3, convert it to unicode assuming it was encoded with UTF-8
        self.vba_code_all_modules = bytes2str(data)
        self.contains_vba_macros = True
        # set type only if parsing succeeds
        self.type = TYPE_TEXT

    def append_subfile(self, filename, data, container=None):
        """
        Create sub-parser for given subfile/data and append to subfiles.
        """
        self.ole_subfiles.append(VBA_Parser(filename, data, container,
                                            relaxed=self.relaxed,
                                            encoding=self.encoding,
                                            disable_pcode=self.disable_pcode))

    def find_vba_projects(self):
        """
        Finds all the VBA projects stored in an OLE file.

        Return None if the file is not OLE but OpenXML.
        Return a list of tuples (vba_root, project_path, dir_path) for each VBA project.
        vba_root is the path of the root OLE storage containing the VBA project,
        including a trailing slash unless it is the root of the OLE file.
        project_path is the path of the OLE stream named "PROJECT" within the VBA project.
        dir_path is the path of the OLE stream named "VBA/dir" within the VBA project.

        If this function returns an empty list for one of the supported formats
        (i.e. Word, Excel, Powerpoint), then the file does not contain VBA macros.

        :return: None if OpenXML file, list of tuples (vba_root, project_path, dir_path)
        for each VBA project found if OLE file
        """
        log.debug('VBA_Parser.find_vba_projects')

        # if the file is not OLE but OpenXML, return None:
        if self.ole_file is None and self.type != TYPE_PPT:
            return None

        # if this method has already been called, return previous result:
        if self.vba_projects is not None:
            return self.vba_projects

        # if this is a ppt file (PowerPoint 97-2003):
        # self.ole_file is None but the ole_subfiles do contain vba_projects
        # (like for OpenXML files).
        if self.type == TYPE_PPT:
            # TODO: so far, this function is never called for PPT files, but
            # if that happens, the information is lost which ole file contains
            # which storage!
            log.warning('Returned info is not complete for PPT types!')
            self.vba_projects = []
            for subfile in self.ole_subfiles:
                self.vba_projects.extend(subfile.find_vba_projects())
            return self.vba_projects

        # Find the VBA project root (different in MS Word, Excel, etc):
        # - Word 97-2003: Macros
        # - Excel 97-2003: _VBA_PROJECT_CUR
        # - PowerPoint 97-2003: PptParser has identified ole_subfiles
        # - Word 2007+: word/vbaProject.bin in zip archive, then the VBA project is the root of vbaProject.bin.
        # - Excel 2007+: xl/vbaProject.bin in zip archive, then same as Word
        # - PowerPoint 2007+: ppt/vbaProject.bin in zip archive, then same as Word
        # - Visio 2007: not supported yet (different file structure)

        # According to MS-OVBA section 2.2.1:
        # - the VBA project root storage MUST contain a VBA storage and a PROJECT stream
        # - The root/VBA storage MUST contain a _VBA_PROJECT stream and a dir stream
        # - all names are case-insensitive

        def check_vba_stream(ole, vba_root, stream_path):
            full_path = vba_root + stream_path
            if ole.exists(full_path) and ole.get_type(full_path) == olefile.STGTY_STREAM:
                log.debug('Found %s stream: %s' % (stream_path, full_path))
                return full_path
            else:
                log.debug('Missing %s stream, this is not a valid VBA project structure' % stream_path)
                return False

        # start with an empty list:
        self.vba_projects = []
        # Look for any storage containing those storage/streams:
        ole = self.ole_file
        for storage in ole.listdir(streams=False, storages=True):
            log.debug('Checking storage %r' % storage)
            # Look for a storage ending with "VBA":
            if storage[-1].upper() == 'VBA':
                log.debug('Found VBA storage: %s' % ('/'.join(storage)))
                vba_root = '/'.join(storage[:-1])
                # Add a trailing slash to vba_root, unless it is the root of the OLE file:
                # (used later to append all the child streams/storages)
                if vba_root != '':
                    vba_root += '/'
                log.debug('Checking vba_root="%s"' % vba_root)

                # Check if the VBA root storage also contains a PROJECT stream:
                project_path = check_vba_stream(ole, vba_root, 'PROJECT')
                if not project_path: continue
                # Check if the VBA root storage also contains a VBA/_VBA_PROJECT stream:
                vba_project_path = check_vba_stream(ole, vba_root, 'VBA/_VBA_PROJECT')
                if not vba_project_path: continue
                # Check if the VBA root storage also contains a VBA/dir stream:
                dir_path = check_vba_stream(ole, vba_root, 'VBA/dir')
                if not dir_path: continue
                # Now we are pretty sure it is a VBA project structure
                log.debug('VBA root storage: "%s"' % vba_root)
                # append the results to the list as a tuple for later use:
                self.vba_projects.append((vba_root, project_path, dir_path))
        return self.vba_projects

    def detect_macros(self):
        """
        Detect the potential presence of VBA or Excel4/XLM macros in the file,
        by calling detect_vba_macros and detect_xlm_macros.
        (if the no_xlm option is set, XLM macros are not checked)

        :return: bool, True if at least one VBA project has been found, False otherwise
        """
        vba = self.detect_vba_macros()
        xlm = False
        if not self.no_xlm:
            xlm = self.detect_xlm_macros()
        return (vba or xlm)

    def detect_vba_macros(self):
        """
        Detect the potential presence of VBA macros in the file, by checking
        if it contains VBA projects. Both OLE and OpenXML files are supported.

        Important: for now, results are accurate only for Word, Excel and PowerPoint

        Note: this method does NOT attempt to check the actual presence or validity
        of VBA macro source code, so there might be false positives.
        It may also detect VBA macros in files embedded within the main file,
        for example an Excel workbook with macros embedded into a Word
        document without macros may be detected, without distinction.

        :return: bool, True if at least one VBA project has been found, False otherwise
        """
        log.debug("detect vba macros")
        # TODO: return None or raise exception if format not supported
        # TODO: return the number of VBA projects found instead of True/False?
        # if this method was already called, return the previous result:
        if self.contains_vba_macros is not None:
            return self.contains_vba_macros
        # if OpenXML/PPT, check all the OLE subfiles:
        if self.ole_file is None:
            for ole_subfile in self.ole_subfiles:
                log.debug("ole subfile {}".format(ole_subfile))
                ole_subfile.no_xlm = self.no_xlm
                if ole_subfile.detect_vba_macros():
                    self.contains_vba_macros = True
                    return True
            # otherwise, no macro found:
            self.contains_vba_macros = False
            return False
        # otherwise it's an OLE file, find VBA projects:
        vba_projects = self.find_vba_projects()
        if len(vba_projects) == 0:
            self.contains_vba_macros = False
        else:
            self.contains_vba_macros = True
        # Also look for VBA code in any stream including orphans
        # (happens in some malformed files)
        ole = self.ole_file
        for sid in xrange(len(ole.direntries)):
            # check if id is already done above:
            log.debug('Checking DirEntry #%d' % sid)
            d = ole.direntries[sid]
            if d is None:
                # this direntry is not part of the tree: either unused or an orphan
                d = ole._load_direntry(sid)
                log.debug('This DirEntry is an orphan or unused')
            if d.entry_type == olefile.STGTY_STREAM:
                # read data
                log.debug('Reading data from stream %r - size: %d bytes' % (d.name, d.size))
                try:
                    data = ole._open(d.isectStart, d.size).read()
                    log.debug('Read %d bytes' % len(data))
                    if len(data) > 200:
                        log.debug('%r...[much more data]...%r' % (data[:100], data[-50:]))
                    else:
                        log.debug(repr(data))
                    if b'Attribut\x00' in data:
                        log.debug('Found VBA compressed code')
                        self.contains_vba_macros = True
                except IOError as exc:
                    if self.relaxed:
                        log.info('Error when reading OLE Stream %r' % d.name)
                        log.debug('Trace:', exc_trace=True)
                    else:
                        raise SubstreamOpenError(self.filename, d.name, exc)
        return self.contains_vba_macros

    def detect_xlm_macros(self):
        """
        Detect the potential presence of Excel 4/XLM macros in the file, by checking
        if it contains a macro worksheet. Both OLE and OpenXML files are supported.
        Only Excel files may contain XLM macros, and also SLK and CSV files.

        If XLMMacroDeobfuscator is available, it will be used. Otherwise plugin_biff
        is used as fallback (plugin_biff only supports OLE files, not XLSX or XLSB)

        :return: bool, True if at least one macro worksheet has been found, False otherwise
        """
        log.debug("detect xlm macros")
        # if this method was already called, return the previous result:
        if self.contains_xlm_macros is not None:
            return self.contains_xlm_macros
        # if this is a SLK file, the analysis was done in open_slk:
        if self.type == TYPE_SLK:
            return self.contains_xlm_macros
        # TODO: check also CSV files for formulas?
        self.xlm_macros = []
        # check if the file is Excel, otherwise return False
        if not self.ftg.is_excel():
            self.contains_xlm_macros = False
            return False
        if XLMDEOBFUSCATOR:
            # XLMMacroDeobfuscator is available, use it:
            # But it only works with files on disk for now
            if not self.file_on_disk:
                log.warning('XLMMacroDeobfuscator only works with files on disk, not in memory. Analysis might be less complete.')
            else:
                try:
                    return self._extract_xlm_xlmdeobf()
                except Exception:
                    log.error('Error when running XLMMacroDeobfuscator')
        # fall back to plugin_biff:
        if self.ole_file is None:
            # TODO: handle OpenXML here
            return False
        return self._extract_xlm_plugin_biff()

    def _extract_xlm_xlmdeobf(self):
        """
        Run XLMMacroDeobfuscator to detect and extract XLM macros
        :return: bool, True if at least one macro worksheet has been found, False otherwise
        """
        log.debug('Calling XLMMacroDeobfuscator to detect and extract XLM macros')
        xlmdeobfuscator.SILENT = True
        # we build the output as a list of strings:
        xlm = ["RAW EXCEL4/XLM MACRO FORMULAS:"]
        # First, extract only formulas without emulation
        result = xlmdeobfuscator.process_file(file=self.filename,
                                           noninteractive=True,
                                           noindent=True,
                                           # output_formula_format='CELL:[[CELL_ADDR]], [[INT-FORMULA]]',
                                           return_deobfuscated=True,
                                           timeout=30,
                                           extract_only=True,
                                           )
        if len(result) == 0:
            # no XLM macro was found
            self.contains_xlm_macros = False
            return False
        xlm += result
        xlm.append('- ' * 38)
        xlm.append('EMULATION - DEOBFUSCATED EXCEL4/XLM MACRO FORMULAS:')
        result = xlmdeobfuscator.process_file(file=self.filename,
                                           noninteractive=True,
                                           noindent=True,
                                           # output_formula_format='CELL:[[CELL_ADDR]], [[INT-FORMULA]]',
                                           return_deobfuscated=True,
                                           timeout=30,
                                           )
        xlm += result
        log.debug(xlm)
        self.xlm_macros = xlm
        self.contains_xlm_macros = True
        return True


    def _extract_xlm_plugin_biff(self):
        """
        Run plugin_biff to detect and extract XLM macros
        :return: bool, True if at least one macro worksheet has been found, False otherwise
        """
        log.debug('_extract_xlm_plugin_biff')
        for excel_stream in ('Workbook', 'Book'):
            if self.ole_file.exists(excel_stream):
                log.debug('Found Excel stream %r' % excel_stream)
                data = self.ole_file.openstream(excel_stream).read()
                log.debug('Running BIFF plugin from oledump')
                try:
                    # starting from plugin_biff 0.0.12, we use the CSV output (-c) instead of -x
                    # biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-x')
                    # First let's get the list of boundsheets, and check if there are Excel 4 macros:
                    biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-o BOUNDSHEET')
                    self.xlm_macros = biff_plugin.Analyze()
                    if "Excel 4.0 macro sheet" in '\n'.join(self.xlm_macros):
                        log.debug('Found XLM macros')
                        # get the list of labels, which may contain the "Auto_Open" trigger
                        biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-o LABEL -r LN')
                        self.xlm_macros += biff_plugin.Analyze()
                        biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-c -r LN')
                        self.xlm_macros += biff_plugin.Analyze()
                        # we run plugin_biff again, this time to search DCONN objects and get their URLs, if any:
                        # ref: https://inquest.net/blog/2020/03/18/Getting-Sneakier-Hidden-Sheets-Data-Connections-and-XLM-Macros
                        biff_plugin = cBIFF(name=[excel_stream], stream=data, options='-o DCONN -s')
                        self.xlm_macros += biff_plugin.Analyze()
                        self.contains_xlm_macros = True
                        return True
                except:
                    log.exception('Error when running oledump.plugin_biff, please report to %s' % URL_OLEVBA_ISSUES)
        self.contains_xlm_macros = False
        return False

    def detect_is_encrypted(self):
        if self.ole_file:
            self.is_encrypted = crypto.is_encrypted(self.ole_file)
        return self.is_encrypted

    def decrypt_file(self, passwords_list=None):
        decrypted_file = None
        if self.detect_is_encrypted():
            passwords = crypto.DEFAULT_PASSWORDS
            if passwords_list and isinstance(passwords_list, list):
                passwords.extend(passwords_list)
            decrypted_file = crypto.decrypt(self.filename, passwords)

        return decrypted_file

    def encode_string(self, unicode_str):
        """
        Encode a unicode string to bytes or str, using the specified encoding
        for the VBA_parser. By default, it will be bytes/UTF-8 on Python 2, and
        a normal unicode string on Python 3.
        :param str unicode_str: string to be encoded
        :return: encoded string
        """
        if self.encoding is None:
            return unicode_str
        else:
            return unicode_str.encode(self.encoding, errors='replace')

    def extract_macros(self):
        """
        Extract and decompress source code for each VBA macro found in the file

        Iterator: yields (filename, stream_path, vba_filename, vba_code) for each VBA macro found
        If the file is OLE, filename is the path of the file.
        If the file is OpenXML, filename is the path of the OLE subfile containing VBA macros
        within the zip archive, e.g. word/vbaProject.bin.
        If the file is PPT, result is as for OpenXML but filename is useless
        """
        log.debug('extract_macros:')
        if self.ole_file is None:
            # This may be either an OpenXML/PPT or a text file:
            if self.type == TYPE_TEXT:
                # This is a text file, yield the full code:
                yield (self.filename, '', self.filename, self.vba_code_all_modules)
            elif self.type == TYPE_SLK:
                if self.xlm_macros:
                    vba_code = ''
                    for line in self.xlm_macros:
                        vba_code += "' " + line + '\n'
                    yield ('xlm_macro', 'xlm_macro', 'xlm_macro.txt', vba_code)
            else:
                # OpenXML/PPT: recursively yield results from each OLE subfile:
                for ole_subfile in self.ole_subfiles:
                    for results in ole_subfile.extract_macros():
                        yield results
                # we also need to yield XLM macros
                if self.xlm_macros:
                    vba_code = ''
                    for line in self.xlm_macros:
                        vba_code += "' " + line + '\n'
                    yield ('xlm_macro', 'xlm_macro', 'xlm_macro.txt', vba_code)
        else:
            # This is an OLE file:
            self.find_vba_projects()
            # set of stream ids
            vba_stream_ids = set()
            for vba_root, project_path, dir_path in self.vba_projects:
                # extract all VBA macros from that VBA root storage:
                # The function _extract_vba may fail on some files (issue #132)
                # TODO: refactor this loop, because if one module fails it stops parsing,
                #  and the error is only logged, not stored for reporting anomalies
                try:
                    for stream_path, vba_filename, vba_code in \
                            _extract_vba(self.ole_file, vba_root, project_path,
                                         dir_path, self.relaxed):
                        # store direntry ids in a set:
                        vba_stream_ids.add(self.ole_file._find(stream_path))
                        yield (self.filename, stream_path, vba_filename, vba_code)
                except Exception as e:
                    log.exception('Error in _extract_vba')
            # Also look for VBA code in any stream including orphans
            # (happens in some malformed files)
            ole = self.ole_file
            for sid in xrange(len(ole.direntries)):
                # check if id is already done above:
                log.debug('Checking DirEntry #%d' % sid)
                if sid in vba_stream_ids:
                    log.debug('Already extracted')
                    continue
                d = ole.direntries[sid]
                if d is None:
                    # this direntry is not part of the tree: either unused or an orphan
                    d = ole._load_direntry(sid)
                    log.debug('This DirEntry is an orphan or unused')
                if d.entry_type == olefile.STGTY_STREAM:
                    # read data
                    log.debug('Reading data from stream %r' % d.name)
                    data = ole._open(d.isectStart, d.size).read()
                    for match in re.finditer(b'\\x00Attribut[^e]', data, flags=re.IGNORECASE):
                        start = match.start() - 3
                        log.debug('Found VBA compressed code at index %X' % start)
                        compressed_code = data[start:]
                        try:
                            vba_code_bytes = decompress_stream(bytearray(compressed_code))
                            # vba_code_bytes is in bytes, we need to convert it to str
                            # but here we don't know the encoding of the VBA project
                            # (for example code page 1252 or 1251), because it's in the
                            # VBA_Project class and if we're here it may be because
                            # the VBA project parsing failed (e.g. issue #593).
                            # So let's convert using cp1252 as a guess
                            # TODO get the actual encoding from the VBA_Project
                            vba_code_str = bytes2str(vba_code_bytes, encoding='cp1252')
                            yield (self.filename, d.name, d.name, vba_code_str)
                        except Exception as exc:
                            # display the exception with full stack trace for debugging
                            log.debug('Error processing stream %r in file %r (%s)' % (d.name, self.filename, exc))
                            log.debug('Traceback:', exc_info=True)
                            # do not raise the error, as it is unlikely to be a compressed macro stream
            if self.xlm_macros:
                vba_code = ''
                for line in self.xlm_macros:
                    vba_code += "' " + line + '\n'
                yield ('xlm_macro', 'xlm_macro', 'xlm_macro.txt', vba_code)
            # Analyse the VBA P-code to detect VBA stomping:
            # If stomping is detected, add a fake VBA module with the P-code as source comments
            # so that VBA_Scanner can find keywords and IOCs in it
            if self.detect_vba_stomping():
                vba_code = ''
                for line in self.pcodedmp_output.splitlines():
                    vba_code += "' " + line + '\n'
                yield ('VBA P-code', 'VBA P-code', 'VBA_P-code.txt', vba_code)


    def extract_all_macros(self):
        """
        Extract and decompress source code for each VBA macro found in the file
        by calling extract_macros(), store the results as a list of tuples
        (filename, stream_path, vba_filename, vba_code) in self.modules.
        See extract_macros for details.
        :returns: list of tuples (filename, stream_path, vba_filename, vba_code)
        """
        if self.modules is None:
            self.modules = []
            for (subfilename, stream_path, vba_filename, vba_code) in self.extract_macros():
                self.modules.append((subfilename, stream_path, vba_filename, vba_code))
        self.nb_macros = len(self.modules)
        return self.modules


    def get_vba_code_all_modules(self):
        """
        Extract the VBA macro source code from all modules, and return it
        as a single string (str) with all modules concatenated.
        If an exception is triggered when decompressing a VBA module, it
        will not be included. The error is logged but the exception is not
        raised further.
        :return: str
        """
        vba_code_all_modules = ''
        for (_, _, _, vba_code) in self.extract_all_macros():
            if not isinstance(vba_code, str):
                log.error('VBA code returned by extract_all_macros is not a string')
            else:
                vba_code_all_modules += vba_code + '\n'
        return vba_code_all_modules


    def analyze_macros(self, show_decoded_strings=False, deobfuscate=False):
        """
        runs extract_macros and analyze the source code of all VBA+XLM macros
        found in the file.
        All results are stored in self.analysis_results.
        If called more than once, simply returns the previous results.

        :return: list of tuples (type, keyword, description)
        (type = 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String' or 'Dridex String')
        """
        # Check if there are VBA or XLM macros:
        if self.detect_macros():
            # if the analysis was already done, avoid doing it twice:
            if self.analysis_results is not None:
                return self.analysis_results
            # variable to merge source code from all modules:
            if self.vba_code_all_modules is None:
                self.vba_code_all_modules = self.get_vba_code_all_modules()
                for (_, _, form_string) in self.extract_form_strings():
                    self.vba_code_all_modules += form_string + '\n'
            # Analyze the whole code at once:
            scanner = VBA_Scanner(self.vba_code_all_modules)
            self.analysis_results = scanner.scan(show_decoded_strings, deobfuscate)
            if self.detect_vba_stomping():
                log.debug('adding VBA stomping to suspicious keywords')
                keyword = 'VBA Stomping'
                description = 'VBA Stomping was detected: the VBA source code and P-code are different, '\
                    'this may have been used to hide malicious code'
                scanner.suspicious_keywords.append((keyword, description))
                scanner.results.append(('Suspicious', keyword, description))
            if self.contains_xlm_macros:
                log.debug('adding XLM macrosheet found to suspicious keywords')
                keyword = 'XLM macro'
                description = 'XLM macro found. It may contain malicious code'
                scanner.suspicious_keywords.append((keyword, description))
                scanner.results.append(('Suspicious', keyword, description))
            # TODO: this has been temporarily disabled
            if self.template_injection_found:
                log.debug('adding Template Injection to suspicious keywords')
                keyword = 'Template Injection'
                description = 'Template injection found. A malicious template could have been uploaded ' \
                    'from a remote location'
                scanner.suspicious_keywords.append((keyword, description))
                scanner.results.append(('Suspicious', keyword, description))
            autoexec, suspicious, iocs, hexstrings, base64strings, dridex, vbastrings = scanner.scan_summary()
            self.nb_autoexec += autoexec
            self.nb_suspicious += suspicious
            self.nb_iocs += iocs
            self.nb_hexstrings += hexstrings
            self.nb_base64strings += base64strings
            self.nb_dridexstrings += dridex
            self.nb_vbastrings += vbastrings

        return self.analysis_results


    def reveal(self):
        # we only want printable strings:
        analysis = self.analyze_macros(show_decoded_strings=False)
        # to avoid replacing short strings contained into longer strings, we sort the analysis results
        # based on the length of the encoded string, in reverse order:
        analysis = sorted(analysis, key=lambda type_decoded_encoded: len(type_decoded_encoded[2]), reverse=True)
        # normally now self.vba_code_all_modules contains source code from all modules
        # Need to collapse long lines:
        deobf_code = vba_collapse_long_lines(self.vba_code_all_modules)
        deobf_code = filter_vba(deobf_code)
        for kw_type, decoded, encoded in analysis:
            if kw_type == 'VBA string':
                #print '%3d occurences: %r => %r' % (deobf_code.count(encoded), encoded, decoded)
                # need to add double quotes around the decoded strings
                # after escaping double-quotes as double-double-quotes for VBA:
                decoded = decoded.replace('"', '""')
                decoded = '"%s"' % decoded
                # if the encoded string is enclosed in parentheses,
                # keep them in the decoded version:
                if encoded.startswith('(') and encoded.endswith(')'):
                    decoded = '(%s)' % decoded
                deobf_code = deobf_code.replace(encoded, decoded)
        # # TODO: there is a bug somewhere which creates double returns '\r\r'
        # deobf_code = deobf_code.replace('\r\r', '\r')
        return deobf_code
        #TODO: repasser l'analyse plusieurs fois si des chaines hex ou base64 sont revelees


    def find_vba_forms(self):
        """
        Finds all the VBA forms stored in an OLE file.

        Return None if the file is not OLE but OpenXML.
        Return a list of tuples (vba_root, project_path, dir_path) for each VBA project.
        vba_root is the path of the root OLE storage containing the VBA project,
        including a trailing slash unless it is the root of the OLE file.
        project_path is the path of the OLE stream named "PROJECT" within the VBA project.
        dir_path is the path of the OLE stream named "VBA/dir" within the VBA project.

        If this function returns an empty list for one of the supported formats
        (i.e. Word, Excel, Powerpoint), then the file does not contain VBA forms.

        :return: None if OpenXML file, list of tuples (vba_root, project_path, dir_path)
        for each VBA project found if OLE file
        """
        log.debug('VBA_Parser.find_vba_forms')

        # if the file is not OLE but OpenXML, return None:
        if self.ole_file is None and self.type != TYPE_PPT:
            return None

        # if this method has already been called, return previous result:
        # if self.vba_projects is not None:
        #     return self.vba_projects

        # According to MS-OFORMS section 2.1.2 Control Streams:
        # - A parent control, that is, a control that can contain embedded controls,
        #   MUST be persisted as a storage that contains multiple streams.
        # - All parent controls MUST contain a FormControl. The FormControl
        #   properties are persisted to a stream (1) as specified in section 2.1.1.2.
        #   The name of this stream (1) MUST be "f".
        # - Embedded controls that cannot themselves contain other embedded
        #   controls are persisted sequentially as FormEmbeddedActiveXControls
        #   to a stream (1) contained in the same storage as the parent control.
        #   The name of this stream (1) MUST be "o".
        # - all names are case-insensitive

        if self.type == TYPE_PPT:
            # TODO: so far, this function is never called for PPT files, but
            # if that happens, the information is lost which ole file contains
            # which storage!
            ole_files = self.ole_subfiles
            log.warning('Returned info is not complete for PPT types!')
        else:
            ole_files = [self.ole_file, ]

        # start with an empty list:
        self.vba_forms = []

        # Loop over ole streams
        for ole in ole_files:
            # Look for any storage containing those storage/streams:
            for storage in ole.listdir(streams=False, storages=True):
                log.debug('Checking storage %r' % storage)
                # Look for two streams named 'o' and 'f':
                o_stream = storage + ['o']
                f_stream = storage + ['f']
                log.debug('Checking if streams %r and %r exist' % (f_stream, o_stream))
                if ole.exists(o_stream) and ole.get_type(o_stream) == olefile.STGTY_STREAM \
                and ole.exists(f_stream) and ole.get_type(f_stream) == olefile.STGTY_STREAM:
                    form_path = '/'.join(storage)
                    log.debug('Found VBA Form: %r' % form_path)
                    self.vba_forms.append(storage)
        return self.vba_forms

    def extract_form_strings(self):
        """
        Extract printable strings from each VBA Form found in the file

        Iterator: yields (filename, stream_path, form_string) for each printable string found in forms
        If the file is OLE, filename is the path of the file.
        If the file is OpenXML, filename is the path of the OLE subfile containing VBA macros
        within the zip archive, e.g. word/vbaProject.bin.
        If the file is PPT, result is as for OpenXML but filename is useless
        Note: form_string is a raw bytes string on Python 2, a unicode str on Python 3
        """
        if self.ole_file is None:
            # This may be either an OpenXML/PPT or a text file:
            if self.type == TYPE_TEXT:
                # This is a text file, return no results:
                return
            else:
                # OpenXML/PPT: recursively yield results from each OLE subfile:
                for ole_subfile in self.ole_subfiles:
                    for results in ole_subfile.extract_form_strings():
                        yield results
        else:
            # This is an OLE file:
            self.find_vba_forms()
            ole = self.ole_file
            for form_storage in self.vba_forms:
                o_stream = form_storage + ['o']
                log.debug('Opening form object stream %r' % '/'.join(o_stream))
                form_data = ole.openstream(o_stream).read()
                # Extract printable strings from the form object stream "o":
                for m in re_printable_string.finditer(form_data):
                    log.debug('Printable string found in form: %r' % m.group())
                    # On Python 3, convert bytes string to unicode str:
                    if PYTHON2:
                        found_str = m.group()
                    else:
                        found_str = m.group().decode('utf8', errors='replace')
                    if found_str != 'Tahoma':
                        yield (self.filename, '/'.join(o_stream), found_str)

    def extract_form_strings_extended(self):
        if self.ole_file is None:
            # This may be either an OpenXML/PPT or a text file:
            if self.type == TYPE_TEXT:
                # This is a text file, return no results:
                return
            else:
                # OpenXML/PPT: recursively yield results from each OLE subfile:
                for ole_subfile in self.ole_subfiles:
                    for results in ole_subfile.extract_form_strings_extended():
                        yield results
        else:
            # This is an OLE file:
            self.find_vba_forms()
            ole = self.ole_file
            for form_storage in self.vba_forms:
                for variable in oleform.extract_OleFormVariables(ole, form_storage):
                    yield (self.filename, '/'.join(form_storage), variable)

    def extract_pcode(self):
        """
        Extract and disassemble the VBA P-code, using pcodedmp

        :return: VBA P-code disassembly
        :rtype: str
        """
        # Text and SLK files cannot be stomped:
        if self.type in (TYPE_SLK, TYPE_TEXT):
            self.pcodedmp_output = ''
            return ''
        # only run it once:
        if self.disable_pcode:
            self.pcodedmp_output = ''
            return ''
        if self.pcodedmp_output is None:
            log.debug('Calling pcodedmp to extract and disassemble the VBA P-code')
            # import pcodedmp here to avoid circular imports:
            try:
                from pcodedmp import pcodedmp
            except Exception as e:
                # This may happen with Pypy, because pcodedmp imports win_unicode_console...
                # TODO: this is a workaround, we just ignore P-code
                # TODO: here we just use log.info, because the word "error" in the output makes some of the tests fail...
                log.info('Exception when importing pcodedmp: {}'.format(e))
                self.pcodedmp_output = ''
                return ''
            # logging is disabled after importing pcodedmp, need to re-enable it
            # This is because pcodedmp imports olevba again :-/
            # TODO: here it works only if logging was enabled, need to change pcodedmp!
            enable_logging()
            # pcodedmp prints all its output to sys.stdout, so we need to capture it so that
            # we can process the results later on.
            # save sys.stdout, then modify it to capture pcodedmp's output:
            # stdout = sys.stdout
            if PYTHON2:
                # on Python 2, console output is bytes
                output = BytesIO()
            else:
                # on Python 3, console output is unicode
                output = StringIO()
            # sys.stdout = output
            # we need to fake an argparser for those two args used by pcodedmp:
            class args:
                disasmOnly = True
                verbose = False
            try:
                # TODO: handle files in memory too
                log.debug('before pcodedmp')
                # TODO: we just ignore pcodedmp errors
                stderr = sys.stderr
                sys.stderr = output
                pcodedmp.processFile(self.filename, args, output_file=output)
                sys.stderr = stderr
                log.debug('after pcodedmp')
            except Exception as e:
                # print('Error while running pcodedmp: {}'.format(e), file=sys.stderr, flush=True)
                # set sys.stdout back to its original value
                # sys.stdout = stdout
                log.exception('Error while running pcodedmp')
            # finally:
            #     # set sys.stdout back to its original value
            #     sys.stdout = stdout
            self.pcodedmp_output = output.getvalue()
            # print(self.pcodedmp_output)
            # log.debug(self.pcodedmp_output)
        return self.pcodedmp_output

    def detect_vba_stomping(self):
        """
        Detect VBA stomping, by comparing the keywords present in the P-code and
        in the VBA source code.

        :return: True if VBA stomping detected, False otherwise
        :rtype: bool
        """
        log.debug('detect_vba_stomping')
        # only run it once:
        if self.vba_stomping_detected is not None:
            return self.vba_stomping_detected
        # Text and SLK files cannot be stomped:
        if self.type in (TYPE_SLK, TYPE_TEXT):
            self.vba_stomping_detected = False
            return False
        # TODO: Files in memory cannot be analysed with pcodedmp yet
        if not self.file_on_disk:
            log.warning('For now, VBA stomping cannot be detected for files in memory')
            self.vba_stomping_detected = False
            return False
        # only run it once:
        if self.vba_stomping_detected is None:
            log.debug('Analysing the P-code to detect VBA stomping')
            self.extract_pcode()
            # print('pcodedmp OK')
            log.debug('pcodedmp OK')
            # process the output to extract keywords, to detect VBA stomping
            keywords = set()
            for line in self.pcodedmp_output.splitlines():
                if line.startswith('\t'):
                    log.debug('P-code: ' + line.strip())
                    tokens = line.split(None, 1)
                    mnemonic = tokens[0]
                    args = ''
                    if len(tokens) == 2:
                        args = tokens[1].strip()
                    # log.debug(repr([mnemonic, args]))
                    # if mnemonic in ('VarDefn',):
                    #     # just add the rest of the line
                    #     keywords.add(args)
                    # if mnemonic == 'FuncDefn':
                    #     # function definition: just strip parentheses
                    #     funcdefn = args.strip('()')
                    #     keywords.add(funcdefn)
                    if mnemonic in ('ArgsCall', 'ArgsLd', 'St', 'Ld', 'MemSt', 'Label'):
                        # sometimes ArgsCall is followed by "(Call)", if so we remove it (issue #489)
                        if args.startswith('(Call) '):
                            args = args[7:]
                        # add 1st argument:
                        name = args.split(None, 1)[0]
                        # sometimes pcodedmp reports names like "id_FFFF", which are not
                        # directly present in the VBA source code
                        # (for example "Me" in VBA appears as id_FFFF in P-code)
                        if not name.startswith('id_'):
                            keywords.add(name)
                    if mnemonic == 'LitStr':
                        # re_string = re.compile(r'\"([^\"]|\"\")*\"')
                        # for match in re_string.finditer(line):
                        #     print('\t' + match.group())
                        # the string is the 2nd argument:
                        s = args.split(None, 1)[1]
                        # tricky issue: when a string contains double quotes inside,
                        # pcodedmp returns a single ", whereas in the VBA source code
                        # it is always a double "".
                        # We have to remove the " around the strings, then double the remaining ",
                        # and put back the " around:
                        if len(s)>=2:
                            assert(s[0]=='"' and s[-1]=='"')
                            s = s[1:-1]
                            s = s.replace('"', '""')
                            s = '"' + s + '"'
                        keywords.add(s)
            log.debug('Keywords extracted from P-code: ' + repr(sorted(keywords)))
            self.vba_stomping_detected = False
            # get all VBA code as one string
            vba_code_all_modules = self.get_vba_code_all_modules()
            for keyword in keywords:
                if keyword not in vba_code_all_modules:
                    log.debug('Keyword {!r} not found in VBA code'.format(keyword))
                    log.debug('VBA STOMPING DETECTED!')
                    self.vba_stomping_detected = True
                    break
            if not self.vba_stomping_detected:
                log.debug('No VBA stomping detected.')
        return self.vba_stomping_detected

    def close(self):
        """
        Close all the open files. This method must be called after usage, if
        the application is opening many files.
        """
        if self.ole_file is None:
            if self.ole_subfiles is not None:
                for ole_subfile in self.ole_subfiles:
                    ole_subfile.close()
        else:
            self.ole_file.close()



class VBA_Parser_CLI(VBA_Parser):
    """
    VBA parser and analyzer, adding methods for the command line interface
    of olevba. (see VBA_Parser)
    """

    def __init__(self, *args, **kwargs):
        """
        Constructor for VBA_Parser_CLI.
        Calls __init__ from VBA_Parser with all arguments --> see doc there
        """
        super(VBA_Parser_CLI, self).__init__(*args, **kwargs)


    def run_analysis(self, show_decoded_strings=False, deobfuscate=False):
        """
        Analyze the provided VBA code, without printing the results (yet)
        All results are stored in self.analysis_results.

        :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
        :return: None
        """
        # print a waiting message only if the output is not redirected to a file:
        if sys.stdout.isatty():
            print('Analysis...\r', end='')
            sys.stdout.flush()
        self.analyze_macros(show_decoded_strings, deobfuscate)


    def print_analysis(self, show_decoded_strings=False, deobfuscate=False):
        """
        print the analysis results in a table

        :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
        :return: None
        """
        results = self.analysis_results
        if results:
            t = tablestream.TableStream(column_width=(10, 20, 45),
                                        header_row=('Type', 'Keyword', 'Description'))
            COLOR_TYPE = {
                'AutoExec': 'yellow',
                'Suspicious': 'red',
                'IOC': 'cyan',
            }
            for kw_type, keyword, description in results:
                # handle non printable strings:
                if not is_printable(keyword):
                    keyword = repr(keyword)
                if not is_printable(description):
                    description = repr(description)
                color_type = COLOR_TYPE.get(kw_type, None)
                t.write_row((kw_type, keyword, description), colors=(color_type, None, None))
            t.close()
            if self.vba_stomping_detected:
                print('VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues')
        else:
            print('No suspicious keyword or IOC found.')

    def print_analysis_json(self, show_decoded_strings=False, deobfuscate=False):
        """
        Analyze the provided VBA code, and return the results in json format

        :param vba_code: str, VBA source code to be analyzed
        :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)

        :return: dict
        """
        # print a waiting message only if the output is not redirected to a file:
        if sys.stdout.isatty():
            print('Analysis...\r', end='')
            sys.stdout.flush()
        return [dict(type=kw_type, keyword=keyword, description=description)
                for kw_type, keyword, description in self.analyze_macros(show_decoded_strings, deobfuscate)]

    def colorize_keywords(self, vba_code):
        """
        Colorize keywords found during the VBA code analysis
        :param vba_code: str, VBA code to be colorized
        :return: str, VBA code including color tags for Colorclass
        """
        results = self.analysis_results
        if results:
            COLOR_TYPE = {
                'AutoExec': 'yellow',
                'Suspicious': 'red',
                'IOC': 'cyan',
            }
            for kw_type, keyword, description in results:
                color_type = COLOR_TYPE.get(kw_type, None)
                if color_type:
                    vba_code = vba_code.replace(keyword, '{auto%s}%s{/%s}' % (color_type, keyword, color_type))
        return vba_code

    def process_file(self, show_decoded_strings=False,
                     display_code=True, hide_attributes=True,
                     vba_code_only=False, show_deobfuscated_code=False,
                     deobfuscate=False, show_pcode=False, no_xlm=False):
        """
        Process a single file

        :param filename: str, path and filename of file on disk, or within the container.
        :param data: bytes, content of the file if it is in a container, None if it is a file on disk.
        :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
        :param display_code: bool, if False VBA source code is not displayed (default True)
        :param global_analysis: bool, if True all modules are merged for a single analysis (default),
                                otherwise each module is analyzed separately (old behaviour)
        :param hide_attributes: bool, if True the first lines starting with "Attribute VB" are hidden (default)
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
        :param show_pcode bool: if True, call pcodedmp to disassemble P-code and display it
        :param no_xlm bool: if True, don't use the BIFF plugin to extract old style XLM macros
        """
        #TODO: replace print by writing to a provided output file (sys.stdout by default)
        # fix conflicting parameters:
        self.no_xlm = no_xlm
        if vba_code_only and not display_code:
            display_code = True
        if self.container:
            display_filename = '%s in %s' % (self.filename, self.container)
        else:
            display_filename = self.filename
        print('=' * 79)
        print('FILE: %s' % display_filename)
        try:
            #TODO: handle olefile errors, when an OLE file is malformed
            print('Type: %s'% self.type)
            if self.detect_macros():
                # run analysis before displaying VBA code, in order to colorize found keywords
                self.run_analysis(show_decoded_strings=show_decoded_strings, deobfuscate=deobfuscate)
                #print 'Contains VBA Macros:'
                for (subfilename, stream_path, vba_filename, vba_code) in self.extract_all_macros():
                    if hide_attributes:
                        # hide attribute lines:
                        vba_code_filtered = filter_vba(vba_code)
                    else:
                        vba_code_filtered = vba_code
                    print('-' * 79)
                    print('VBA MACRO %s ' % vba_filename)
                    print('in file: %s - OLE stream: %s' % (subfilename, repr(stream_path)))
                    if display_code:
                        print('- ' * 39)
                        # detect empty macros:
                        if vba_code_filtered.strip() == '':
                            print('(empty macro)')
                        else:
                            # check if the VBA code contains special characters such as backspace (issue #358)
                            if '\x08' in vba_code_filtered:
                                log.warning('The VBA code contains special characters such as backspace, that may be used for obfuscation.')
                                if sys.stdout.isatty():
                                    # if the standard output is the console, we'll display colors
                                    backspace = colorclass.Color(b'{autored}\\x08{/red}')
                                else:
                                    backspace = '\\x08'
                                # replace backspace by "\x08" for display
                                vba_code_filtered = vba_code_filtered.replace('\x08', backspace)
                            try:
                                # Colorize the interesting keywords in the output:
                                # (unless the output is redirected to a file)
                                if sys.stdout.isatty():
                                    vba_code_filtered = colorclass.Color(self.colorize_keywords(vba_code_filtered))
                            except UnicodeError:
                                # TODO better handling of Unicode
                                log.error('Unicode conversion to be fixed before colorizing the output')
                            print(vba_code_filtered)
                for (subfilename, stream_path, form_string) in self.extract_form_strings():
                    if form_string is not None:
                        print('-' * 79)
                        print('VBA FORM STRING IN %r - OLE stream: %r' % (subfilename, stream_path))
                        print('- ' * 39)
                        print(form_string)
                try:
                    for (subfilename, stream_path, form_variables) in self.extract_form_strings_extended():
                        if form_variables is not None:
                            print('-' * 79)
                            print('VBA FORM Variable "%s" IN %r - OLE stream: %r' % (form_variables['name'], subfilename, stream_path))
                            print('- ' * 39)
                            print(str(form_variables['value']))
                except Exception as exc:
                    # display the exception with full stack trace for debugging
                    log.info('Error parsing form: %s' % exc)
                    log.debug('Traceback:', exc_info=True)
                if show_pcode:
                    print('-' * 79)
                    print('P-CODE disassembly:')
                    pcode = self.extract_pcode()
                    print(pcode)
                # if self.type == TYPE_SLK:
                #     # TODO: clean up this code
                #     slk_output = self.vba_code_all_modules
                #     try:
                #         # Colorize the interesting keywords in the output:
                #         # (unless the output is redirected to a file)
                #         if sys.stdout.isatty():
                #             slk_output = colorclass.Color(self.colorize_keywords(slk_output))
                #     except UnicodeError:
                #         # TODO better handling of Unicode
                #         log.debug('Unicode conversion to be fixed before colorizing the output')
                #     print(slk_output)

                if not vba_code_only:
                    # analyse the code from all modules at once:
                    self.print_analysis(show_decoded_strings, deobfuscate)
                if show_deobfuscated_code:
                    print('MACRO SOURCE CODE WITH DEOBFUSCATED VBA STRINGS (EXPERIMENTAL):\n\n')
                    print(self.reveal())
            else:
                print('No VBA or XLM macros found.')
        except OlevbaBaseException:
            raise
        except Exception as exc:
            # display the exception with full stack trace for debugging
            log.info('Error processing file %s (%s)' % (self.filename, exc))
            traceback.print_exc()
            log.debug('Traceback:', exc_info=True)
            raise ProcessingError(self.filename, exc)
        print('')


    def process_file_json(self, show_decoded_strings=False,
                          display_code=True, hide_attributes=True,
                          vba_code_only=False, show_deobfuscated_code=False,
                          deobfuscate=False, show_pcode=False, no_xlm=False):
        """
        Process a single file

        every "show" or "print" here is to be translated as "add to json"

        :param filename: str, path and filename of file on disk, or within the container.
        :param data: bytes, content of the file if it is in a container, None if it is a file on disk.
        :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
        :param display_code: bool, if False VBA source code is not displayed (default True)
        :param global_analysis: bool, if True all modules are merged for a single analysis (default),
                                otherwise each module is analyzed separately (old behaviour)
        :param hide_attributes: bool, if True the first lines starting with "Attribute VB" are hidden (default)
        :param show_deobfuscated_code: bool, if True add deobfuscated code to result
        :param deobfuscate: bool, if True attempt to deobfuscate VBA expressions (slow)
        :param show_pcode: bool, if True add extracted pcode to result
        """
        #TODO: fix conflicting parameters (?)

        self.no_xlm = no_xlm
        if vba_code_only and not display_code:
            display_code = True

        result = {}

        if self.container:
            result['container'] = self.container
        else:
            result['container'] = None
        result['file'] = self.filename
        result['json_conversion_successful'] = False
        result['analysis'] = None
        result['code_deobfuscated'] = None
        result['do_deobfuscate'] = deobfuscate
        result['show_pcode'] = show_pcode

        try:
            #TODO: handle olefile errors, when an OLE file is malformed
            result['type'] = self.type
            macros = []
            if self.detect_macros():
                for (subfilename, stream_path, vba_filename, vba_code) in self.extract_all_macros():
                    curr_macro = {}
                    if hide_attributes:
                        # hide attribute lines:
                        vba_code_filtered = filter_vba(vba_code)
                    else:
                        vba_code_filtered = vba_code

                    curr_macro['vba_filename'] = vba_filename
                    curr_macro['subfilename'] = subfilename
                    curr_macro['ole_stream'] = stream_path
                    if display_code:
                        curr_macro['code'] = vba_code_filtered.strip()
                    else:
                        curr_macro['code'] = None
                    macros.append(curr_macro)
                if not vba_code_only:
                    # analyse the code from all modules at once:
                    result['analysis'] = self.print_analysis_json(show_decoded_strings,
                                                                  deobfuscate)
                if show_deobfuscated_code:
                    result['code_deobfuscated'] = self.reveal()
                if show_pcode:
                    result['pcode'] = self.extract_pcode()
            result['macros'] = macros
            result['json_conversion_successful'] = True
        except Exception as exc:
            # display the exception with full stack trace for debugging
            log.info('Error processing file %s (%s)' % (self.filename, exc))
            log.debug('Traceback:', exc_info=True)
            raise ProcessingError(self.filename, exc)

        return result


    def process_file_triage(self, show_decoded_strings=False, deobfuscate=False, no_xlm=False):
        """
        Process a file in triage mode, showing only summary results on one line.
        """
        #TODO: replace print by writing to a provided output file (sys.stdout by default)
        try:
            #TODO: handle olefile errors, when an OLE file is malformed
            if self.detect_macros():
                # print a waiting message only if the output is not redirected to a file:
                if sys.stdout.isatty():
                    print('Analysis...\r', end='')
                    sys.stdout.flush()
                self.analyze_macros(show_decoded_strings=show_decoded_strings,
                                    deobfuscate=deobfuscate)
            flags = TYPE2TAG[self.type]
            macros = autoexec = suspicious = iocs = hexstrings = base64obf = dridex = vba_obf = '-'
            if self.contains_vba_macros: macros = 'M'
            if self.nb_autoexec: autoexec = 'A'
            if self.nb_suspicious: suspicious = 'S'
            if self.nb_iocs: iocs = 'I'
            if self.nb_hexstrings: hexstrings = 'H'
            if self.nb_base64strings: base64obf = 'B'
            if self.nb_dridexstrings: dridex = 'D'
            if self.nb_vbastrings: vba_obf = 'V'
            flags += '%s%s%s%s%s%s%s%s' % (macros, autoexec, suspicious, iocs, hexstrings,
                                         base64obf, dridex, vba_obf)

            line = '%-12s %s' % (flags, self.filename)
            print(line)
        except Exception as exc:
            # display the exception with full stack trace for debugging only
            log.debug('Error processing file %s (%s)' % (self.filename, exc),
                      exc_info=True)
            raise ProcessingError(self.filename, exc)


#=== MAIN =====================================================================

def parse_args(cmd_line_args=None):
    """ parse command line arguments (given ones or per default sys.argv) """

    DEFAULT_LOG_LEVEL = "warning" # Default log level

    usage = 'usage: olevba [options] <filename> [filename2 ...]'
    parser = argparse.ArgumentParser(usage=usage)
    parser.add_argument('filenames', nargs='*', help='Files to analyze')
    # parser.add_argument('-o', '--outfile', dest='outfile',
    #     help='output file')
    # parser.add_argument('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_argument("-r", action="store_true", dest="recursive",
                        help='find files recursively in subdirectories.')
    parser.add_argument("-z", "--zip", dest='zip_password', type=str,
                        default=None,
                        help='if the file is a zip archive, open all files '
                             'from it, using the provided password.')
    parser.add_argument("-p", "--password", type=str, action='append',
                        default=[],
                        help='if encrypted office files are encountered, try '
                             'decryption with this password. May be repeated.')
    parser.add_argument("-f", "--zipfname", dest='zip_fname', type=str,
                        default='*',
                        help='if the file is a zip archive, file(s) to be '
                             'opened within the zip. Wildcards * and ? are '
                             'supported. (default: %(default)s)')
    modes = parser.add_argument_group(title='Output mode (mutually exclusive)')
    modes.add_argument("-t", '--triage', action="store_const",
                       dest="output_mode", const='triage',
                       default='unspecified',
                       help='triage mode, display results as a summary table '
                            '(default for multiple files)')
    modes.add_argument("-d", '--detailed', action="store_const",
                       dest="output_mode", const='detailed',
                       default='unspecified',
                       help='detailed mode, display full results (default for '
                            'single file)')
    modes.add_argument("-j", '--json', action="store_const",
                       dest="output_mode", const='json', default='unspecified',
                       help='json mode, detailed in json format '
                            '(never default)')
    parser.add_argument("-a", '--analysis', action="store_false",
                        dest="display_code", default=True,
                        help='display only analysis results, not the macro '
                             'source code')
    parser.add_argument("-c", '--code', action="store_true",
                        dest="vba_code_only", default=False,
                        help='display only VBA source code, do not analyze it')
    parser.add_argument("--decode", action="store_true",
                        dest="show_decoded_strings",
                        help='display all the obfuscated strings with their '
                             'decoded content (Hex, Base64, StrReverse, '
                             'Dridex, VBA).')
    parser.add_argument("--attr", action="store_false", dest="hide_attributes",
                        default=True,
                        help='display the attribute lines at the beginning of '
                             'VBA source code')
    parser.add_argument("--reveal", action="store_true",
                        dest="show_deobfuscated_code",
                        help='display the macro source code after replacing '
                             'all the obfuscated strings by their decoded '
                             'content.')
    parser.add_argument('-l', '--loglevel', dest="loglevel", action="store",
                        default=DEFAULT_LOG_LEVEL,
                        help='logging level debug/info/warning/error/critical '
                             '(default=%(default)s)')
    parser.add_argument('--deobf', dest="deobfuscate", action="store_true",
                        default=False,
                        help="Attempt to deobfuscate VBA expressions (slow)")
    # TODO: --relaxed is enabled temporarily until a solution to issue #593 is found
    parser.add_argument('--relaxed', dest="relaxed", action="store_true",
                        default=True,
                        help='Do not raise errors if opening of substream '
                             'fails (this option is now deprecated, enabled by default)')
    parser.add_argument('--show-pcode', dest="show_pcode", action="store_true",
                        default=False,
                        help="Show disassembled P-code (using pcodedmp)")
    parser.add_argument('--no-pcode', action='store_true',
                        help='Disable extraction and analysis of pcode')
    parser.add_argument('--no-xlm', dest="no_xlm", action="store_true", default=False,
                            help="Do not extract XLM Excel macros. This may speed up analysis of large files.")

    options = parser.parse_args(cmd_line_args)

    # Print help if no arguments are passed
    if len(options.filenames) == 0:
        # print banner with version
        python_version = '%d.%d.%d' % sys.version_info[0:3]
        print('olevba %s on Python %s - http://decalage.info/python/oletools' %
              (__version__, python_version))
        print(__doc__)
        parser.print_help()
        sys.exit(RETURN_WRONG_ARGS)

    if options.show_pcode and options.no_pcode:
        parser.error('You cannot combine options --no-pcode and --show-pcode')

    return options


def process_file(filename, data, container, options, crypto_nesting=0):
    """
    Part of main function that processes a single file.
    This is meant to be used only for the command-line interface of olevba

    This handles exceptions and encryption.

    Returns a single code summarizing the status of processing of this file
    """
    try:
        vba_parser = None

        # Open the file
        vba_parser = VBA_Parser_CLI(filename, data=data, container=container,
                                    relaxed=options.relaxed,
                                    disable_pcode=options.no_pcode)

        if options.output_mode == 'detailed':
            # fully detailed output
            vba_parser.process_file(show_decoded_strings=options.show_decoded_strings,
                         display_code=options.display_code,
                         hide_attributes=options.hide_attributes, vba_code_only=options.vba_code_only,
                         show_deobfuscated_code=options.show_deobfuscated_code,
                         deobfuscate=options.deobfuscate, show_pcode=options.show_pcode,
                                    no_xlm=options.no_xlm)
        elif options.output_mode == 'triage':
            # summarized output for triage:
            vba_parser.process_file_triage(show_decoded_strings=options.show_decoded_strings,
                                           deobfuscate=options.deobfuscate, no_xlm=options.no_xlm)
        elif options.output_mode == 'json':
            print_json(
                vba_parser.process_file_json(show_decoded_strings=options.show_decoded_strings,
                         display_code=options.display_code,
                         hide_attributes=options.hide_attributes, vba_code_only=options.vba_code_only,
                         show_deobfuscated_code=options.show_deobfuscated_code,
                         deobfuscate=options.deobfuscate, show_pcode=options.show_pcode,
                                             no_xlm=options.no_xlm))
        else:  # (should be impossible)
            raise ValueError('unexpected output mode: "{0}"!'.format(options.output_mode))
        vba_parser.close()

        # even if processing succeeds, file might still be encrypted
        log.debug('Checking for encryption (normal)')
        if not crypto.is_encrypted(filename):
            log.debug('no encryption detected')
            return RETURN_OK
    except Exception as exc:
        log.debug('Caught exception:', exc_info=True)
        if vba_parser:
            vba_parser.close()

        log.debug('Checking for encryption (after exception)')
        if crypto.is_encrypted(filename):
            pass   # deal with this below
        else:
            if isinstance(exc, (SubstreamOpenError, UnexpectedDataError)):
                if options.output_mode in ('triage', 'unspecified'):
                    print('%-12s %s - Error opening substream or uenxpected ' \
                          'content' % ('?', filename))
                elif options.output_mode == 'json':
                    print_json(file=filename, type='error',
                               error=type(exc).__name__, message=str(exc))
                else:
                    log.exception('Error opening substream or unexpected '
                                  'content in %s' % filename)
                return RETURN_OPEN_ERROR
            elif isinstance(exc, FileOpenError):
                if options.output_mode in ('triage', 'unspecified'):
                    print('%-12s %s - File format not supported' % ('?', filename))
                elif options.output_mode == 'json':
                    print_json(file=filename, type='error',
                               error=type(exc).__name__, message=str(exc))
                else:
                    log.exception('Failed to open %s -- probably not supported!' % filename)
                return RETURN_OPEN_ERROR
            elif isinstance(exc, ProcessingError):
                if options.output_mode in ('triage', 'unspecified'):
                    print('%-12s %s - %s' % ('!ERROR', filename, exc.orig_exc))
                elif options.output_mode == 'json':
                    print_json(file=filename, type='error',
                               error=type(exc).__name__,
                               message=str(exc.orig_exc))
                else:
                    log.exception('Error processing file %s (%s)!'
                                  % (filename, exc.orig_exc))
                return RETURN_PARSE_ERROR
            else:
                raise    # let caller deal with this

    # we reach this point only if file is encrypted
    # check if this is an encrypted file in an encrypted file in an ...
    if crypto_nesting >= crypto.MAX_NESTING_DEPTH:
        raise crypto.MaxCryptoNestingReached(crypto_nesting, filename)

    decrypted_file = None
    try:
        log.debug('Checking encryption passwords {}'.format(options.password))
        passwords = options.password + crypto.DEFAULT_PASSWORDS
        decrypted_file = crypto.decrypt(filename, passwords)
        if not decrypted_file:
            log.error('Decrypt failed, run with debug output to get details')
            raise crypto.WrongEncryptionPassword(filename)
        log.info('Working on decrypted file')
        return process_file(decrypted_file, data, container or filename,
                            options, crypto_nesting+1)
    finally:     # clean up
        try:
            log.debug('Removing crypt temp file {}'.format(decrypted_file))
            os.unlink(decrypted_file)
        except Exception:   # e.g. file does not exist or is None
            pass
    # no idea what to return now
    raise Exception('Programming error -- should never have reached this!')


def main(cmd_line_args=None):
    """
    Main function, called when olevba is run from the command line

    Optional argument: command line arguments to be forwarded to ArgumentParser
    in process_args. Per default (cmd_line_args=None), sys.argv is used. Option
    mainly added for unit-testing
    """
    options = parse_args(cmd_line_args)

    # enable logging in the modules (for json, this prints the opening '['):
    log_helper.enable_logging(options.output_mode=='json', options.loglevel,
                              other_logger_has_first_line=True)

    # provide info about tool and its version
    if options.output_mode == 'json':
        # print first json entry with meta info
        print_json(script_name='olevba', version=__version__,
                   python_version=sys.version_info[0:3],
                   url='http://decalage.info/python/oletools',
                   type='MetaInformation', _json_is_first=True)
    else:
        # print banner with version
        python_version = '%d.%d.%d' % sys.version_info[0:3]
        print('olevba %s on Python %s - http://decalage.info/python/oletools' %
              (__version__, python_version))

    # with the option --reveal, make sure --deobf is also enabled:
    if options.show_deobfuscated_code and not options.deobfuscate:
        log.debug('set --deobf because --reveal was set')
        options.deobfuscate = True

    # gather info on all files that must be processed
    # ignore directory names stored in zip files:
    all_input_info = tuple((container, filename, data) for
                           container, filename, data in xglob.iter_files(
                               options.filenames, recursive=options.recursive,
                               zip_password=options.zip_password,
                               zip_fname=options.zip_fname)
                           if not (container and filename.endswith('/')))

    # specify output mode if options -t, -d and -j were not specified
    if options.output_mode == 'unspecified':
        if len(all_input_info) == 1:
            options.output_mode = 'detailed'
        else:
            options.output_mode = 'triage'

    if options.output_mode == 'triage':
        if options.show_deobfuscated_code:
            log.debug('ignoring option --reveal in triage output mode')
        if options.show_pcode:
            log.debug('ignoring option --show-pcode in triage output mode')

    # Column headers for triage mode
    if options.output_mode == 'triage':
        print('%-12s %-65s' % ('Flags', 'Filename'))
        print('%-12s %-65s' % ('-' * 11, '-' * 65))

    previous_container = None
    count = 0
    container = filename = data = None
    return_code = RETURN_OK
    try:
        for container, filename, data in all_input_info:
            # handle errors from xglob
            if isinstance(data, Exception):
                if isinstance(data, PathNotFoundException):
                    if options.output_mode == 'triage':
                        print('%-12s %s - File not found' % ('?', filename))
                    elif options.output_mode != 'json':
                        log.error('Given path %r does not exist!' % filename)
                    return_code = RETURN_FILE_NOT_FOUND if return_code == 0 \
                                                    else RETURN_SEVERAL_ERRS
                else:
                    if options.output_mode == 'triage':
                        print('%-12s %s - Failed to read from zip file %s' % ('?', filename, container))
                    elif options.output_mode != 'json':
                        log.error('Exception opening/reading %r from zip file %r: %s'
                                      % (filename, container, data))
                    return_code = RETURN_XGLOB_ERR if return_code == 0 \
                                                    else RETURN_SEVERAL_ERRS
                if options.output_mode == 'json':
                    print_json(file=filename, type='error',
                               error=type(data).__name__, message=str(data))
                continue

            if options.output_mode == 'triage':
                # print container name when it changes:
                if container != previous_container:
                    if container is not None:
                        print('\nFiles in %s:' % container)
                    previous_container = container

            # process the file, handling errors and encryption
            curr_return_code = process_file(filename, data, container, options)
            count += 1

            # adjust overall return code
            if curr_return_code == RETURN_OK:
                continue                    # do not modify overall return code
            if return_code == RETURN_OK:
                return_code = curr_return_code      # first error return code
            else:
                return_code = RETURN_SEVERAL_ERRS   # several errors

        if options.output_mode == 'triage':
            print('\n(Flags: OpX=OpenXML, XML=Word2003XML, FlX=FlatOPC XML, MHT=MHTML, TXT=Text, M=Macros, ' \
                  'A=Auto-executable, S=Suspicious keywords, I=IOCs, H=Hex strings, ' \
                  'B=Base64 strings, D=Dridex strings, V=VBA strings, ?=Unknown)\n')

    except crypto.CryptoErrorBase as exc:
        log.exception('Problems with encryption in main: {}'.format(exc),
                      exc_info=True)
        if return_code == RETURN_OK:
            return_code = RETURN_ENCRYPTED
        else:
            return_code == RETURN_SEVERAL_ERRS
    except Exception as exc:
        # some unexpected error, maybe some of the types caught in except clauses
        # above were not sufficient. This is very bad, so log complete trace at exception level
        # and do not care about output mode
        log.exception('Unhandled exception in main: %s' % exc, exc_info=True)
        return_code = RETURN_UNEXPECTED    # even if there were others before -- this is more important
        # TODO: print msg with URL to report issues (except in JSON mode)

    # done. exit
    log.debug('will exit now with code %s' % return_code)
    log_helper.end_logging()
    sys.exit(return_code)

if __name__ == '__main__':
    main()

# This was coded while listening to "Dust" from I Love You But I've Chosen Darkness
