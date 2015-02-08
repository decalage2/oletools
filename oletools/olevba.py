#!/usr/bin/env python
"""
olevba.py

olevba is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to extract VBA Macro code in clear text.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 2007+ (.pptm, .ppsm)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

olevba is part of the python-oletools package:
http://www.decalage.info/python/oletools

olevba is based on source code from officeparser by John William Davison
https://github.com/unixfreak0037/officeparser
"""

#=== LICENSE ==================================================================

# olevba is copyright (c) 2014-2015 Philippe Lagadec (http://www.decalage.info)
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

#------------------------------------------------------------------------------
# CHANGELOG:
# 2014-08-05 v0.01 PL: - first version based on officeparser code
# 2014-08-14 v0.02 PL: - fixed bugs in code, added license from officeparser
# 2014-08-15       PL: - fixed incorrect value check in PROJECTHELPFILEPATH Record
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

__version__ = '0.24'

#------------------------------------------------------------------------------
# TODO:
# + do not use logging, but a provided logger (null logger by default)
# + setup logging (common with other oletools)
# + add xor bruteforcing like bbharvest
# + add chr() decoding

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
# - add support for PowerPoint macros (see libclamav, libgsf), use oledump heuristic?
# - check VBA macros in Visio, Access, Project, etc
# - extract_macros: convert to a class, split long function into smaller methods
# - extract_macros: read bytes from stream file objects instead of strings
# - extract_macros: use combined struct.unpack instead of many calls

#------------------------------------------------------------------------------
# REFERENCES:
# - [MS-OVBA]: Microsoft Office VBA File Format Structure
#   http://msdn.microsoft.com/en-us/library/office/cc313094%28v=office.12%29.aspx
# - officeparser: https://github.com/unixfreak0037/officeparser


#--- IMPORTS ------------------------------------------------------------------

import sys, logging
import struct
import cStringIO
import math
import zipfile
import re
import optparse
import os.path
import binascii
import base64
import traceback

import thirdparty.olefile as olefile
from thirdparty.prettytable import prettytable
from thirdparty.xglob import xglob

#--- CONSTANTS ----------------------------------------------------------------

TYPE_OLE     = 'OLE'
TYPE_OpenXML = 'OpenXML'

MODULE_EXTENSION = "bas"
CLASS_EXTENSION = "cls"
FORM_EXTENSION = "frm"

# Keywords to detect auto-executable macros
AUTOEXEC_KEYWORDS = {
    # MS Word:
    'Runs when the Word document is opened':
        ('AutoExec', 'AutoOpen', 'Document_Open', 'DocumentOpen'),
    'Runs when the Word document is closed':
        ('AutoExit', 'AutoClose', 'Document_Close', 'DocumentBeforeClose'),
    'Runs when the Word document is modified':
        ('DocumentChange',),
    'Runs when a new Word document is created':
        ('AutoNew', 'Document_New', 'NewDocument'),

    # MS Excel:
    'Runs when the Excel Workbook is opened':
        ('Auto_Open', 'Workbook_Open'),
    'Runs when the Excel Workbook is closed':
        ('Auto_Close', 'Workbook_Close'),

    #TODO: full list in MS specs??
}

# Suspicious Keywords that may be used by malware
# See VBA language reference: http://msdn.microsoft.com/en-us/library/office/jj692818%28v=office.15%29.aspx
SUSPICIOUS_KEYWORDS = {
    #TODO: use regex to support variable whitespaces
    'May read system environment variables':
        ('Environ',),
    'May open a file':
        ('Open',),
    'May write to a file (if combined with Open)':
        #TODO: regex to find Open+Write on same line
        ('Write', 'Put', 'Output', 'Print #'),
    'May read or write a binary file (if combined with Open)':
        #TODO: regex to find Open+Binary on same line
        ('Binary',),
    'May copy a file':
        ('FileCopy', 'CopyFile'),
        #FileCopy: http://msdn.microsoft.com/en-us/library/office/gg264390%28v=office.15%29.aspx
        #CopyFile: http://msdn.microsoft.com/en-us/library/office/gg264089%28v=office.15%29.aspx
    'May delete a file':
        ('Kill',),
    'May create a text file':
        ('CreateTextFile','ADODB.Stream', 'WriteText', 'SaveToFile'),
        #CreateTextFile: http://msdn.microsoft.com/en-us/library/office/gg264617%28v=office.15%29.aspx
        #ADODB.Stream sample: http://pastebin.com/Z4TMyuq6
    'May run an executable file or a system command':
        ('Shell', 'vbNormal', 'vbNormalFocus', 'vbHide', 'vbMinimizedFocus', 'vbMaximizedFocus', 'vbNormalNoFocus',
         'vbMinimizedNoFocus', 'WScript.Shell', 'Run'),
        #Shell: http://msdn.microsoft.com/en-us/library/office/gg278437%28v=office.15%29.aspx
        #WScript.Shell+Run sample: http://pastebin.com/Z4TMyuq6
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
    'May run an application (if combined with CreateObject)':
        ('Shell.Application',),
    'May enumerate application windows (if combined with Shell.Application object)':
        ('Windows', 'FindWindow'),
    'May run code from a DLL':
        #TODO: regex to find declare+lib on same line
        ('Lib',),
    'May download files from the Internet':
        #TODO: regex to find urlmon+URLDownloadToFileA on same line
        ('URLDownloadToFileA', 'Msxml2.XMLHTTP', 'Microsoft.XMLHTTP'),
    'May control another application by simulating user keystrokes':
        ('SendKeys', 'AppActivate'),
        #SendKeys: http://msdn.microsoft.com/en-us/library/office/gg278655%28v=office.15%29.aspx
    'May attempt to obfuscate malicious function calls':
        ('CallByName',),
        #CallByName: http://msdn.microsoft.com/en-us/library/office/gg278760%28v=office.15%29.aspx
    'May attempt to obfuscate specific strings':
        #TODO: regex to find several Chr*, not just one
        ('Chr', 'ChrB', 'ChrW', 'StrReverse', 'Xor'),
        #Chr: http://msdn.microsoft.com/en-us/library/office/gg264465%28v=office.15%29.aspx
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
IPv4 = r'(?:'+NUMBER_0_255+r'\.){3}'+NUMBER_0_255
# IPv4 must come before the DNS name because it is more specific
SERVER = r'(?:' + IPv4 + '|' + DNS_NAME + ')'
PORT = r'(?:\:[0-9]{1,5})?'
SERVER_PORT = SERVER + PORT
URL_PATH = r'(?:/[a-zA-Z0-9\-\._\?\,\'/\\\+&%\$#\=~]*)?' # [^\.\,\)\(\s"]
URL_RE = SCHEME + r'\://' + SERVER_PORT + URL_PATH
re_url = re.compile(URL_RE)


# Patterns to be extracted (IP addresses, URLs, etc)
# From patterns.py in balbuzard
RE_PATTERNS = (
    ('URL', re.compile(URL_RE)),
    ('IPv4 address', re.compile(IPv4)),
    ('E-mail address', re.compile(r'(?i)\b[A-Z0-9._%+-]+@'+SERVER+'\b')),
    # ('Domain name', re.compile(r'(?=^.{1,254}$)(^(?:(?!\d+\.|-)[a-zA-Z0-9_\-]{1,63}(?<!-)\.?)+(?:[a-zA-Z]{2,})$)')),
    # Executable file name with known extensions (except .com which is present in many URLs, and .application):
    ("Executable file name", re.compile(r"(?i)\b\w+\.(EXE|PIF|GADGET|MSI|MSP|MSC|VBS|VBE|VB|JSE|JS|WSF|WSC|WSH|WS|BAT|CMD|DLL|SCR|HTA|CPL|CLASS|JAR|PS1XML|PS1|PS2XML|PS2|PSC1|PSC2|SCF|LNK|INF|REG)\b")),
    # Sources: http://www.howtogeek.com/137270/50-file-extensions-that-are-potentially-dangerous-on-windows/
    #TODO: https://support.office.com/en-us/article/Blocked-attachments-in-Outlook-3811cddc-17c3-4279-a30c-060ba0207372#__attachment_file_types
    #('Hex string', re.compile(r'(?:[0-9A-Fa-f]{2}){4,}')),
    )

# regex to detect strings encoded in hexadecimal
re_hex_string = re.compile(r'(?:[0-9A-Fa-f]{2}){4,}')

# regex to detect strings encoded in base64
#re_base64_string = re.compile(r'"(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?"')
# better version from balbuzard, less false positives:
re_base64_string = re.compile(r'"(?:[A-Za-z0-9+/]{4}){1,}(?:[A-Za-z0-9+/]{2}[AEIMQUYcgkosw048]=|[A-Za-z0-9+/][AQgw]==)?"')
# white list of common strings matching the base64 regex, but which are not base64 strings (all lowercase):
BASE64_WHITELIST = set(['thisdocument', 'thisworkbook', 'test', 'temp', 'http', 'open', 'exit'])

# regex to detect strings encoded with a specific Dridex algorithm
# (see https://github.com/JamesHabben/MalwareStuff)
re_dridex_string = re.compile(r'"[0-9A-Za-z]{20,}"')
# regex to check that it is not just a hex string:
re_nothex_check = re.compile(r'[G-Zg-z]')

#--- FUNCTIONS ----------------------------------------------------------------

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


def decompress_stream (compressed_container):
    """
    Decompress a stream according to MS-OVBA section 2.4.1

    compressed_container: string compressed according to the MS-OVBA 2.4.1.3.6 Compression algorithm
    return the decompressed container as a string (bytes)
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

    decompressed_container = ''  # result
    compressed_current = 0

    sig_byte = ord(compressed_container[compressed_current])
    if sig_byte != 0x01:
        raise ValueError('invalid signature byte {0:02X}'.format(sig_byte))

    compressed_current += 1

    #NOTE: the definition of CompressedRecordEnd is ambiguous. Here we assume that
    # CompressedRecordEnd = len(compressed_container)
    while compressed_current < len(compressed_container):
        # 2.4.1.1.5
        compressed_chunk_start = compressed_current
        # chunk header = first 16 bits
        compressed_chunk_header = struct.unpack("<H", compressed_container[compressed_chunk_start:compressed_chunk_start + 2])[0]
        # chunk size = 12 first bits of header + 3
        chunk_size = (compressed_chunk_header & 0x0FFF) + 3
        # chunk signature = 3 next bits - should always be 0b011
        chunk_signature = (compressed_chunk_header >> 12) & 0x07
        if chunk_signature != 0b011:
            raise ValueError('Invalid CompressedChunkSignature in VBA compressed stream')
        # chunk flag = next bit - 1 == compressed, 0 == uncompressed
        chunk_flag = (compressed_chunk_header >> 15) & 0x01
        logging.debug("chunk size = {0}, compressed flag = {1}".format(chunk_size, chunk_flag))

        #MS-OVBA 2.4.1.3.12: the maximum size of a chunk including its header is 4098 bytes (header 2 + data 4096)
        # The minimum size is 3 bytes
        # NOTE: there seems to be a typo in MS-OVBA, the check should be with 4098, not 4095 (which is the max value
        # in chunk header before adding 3.
        # Also the first test is not useful since a 12 bits value cannot be larger than 4095.
        if chunk_flag == 1 and chunk_size > 4098:
            raise ValueError('CompressedChunkSize > 4098 but CompressedChunkFlag == 1')
        if chunk_flag == 0 and chunk_size != 4098:
            raise ValueError('CompressedChunkSize != 4098 but CompressedChunkFlag == 0')

        # check if chunk_size goes beyond the compressed data, instead of silently cutting it:
        #TODO: raise an exception?
        if compressed_chunk_start + chunk_size > len(compressed_container):
            logging.warning('Chunk size is larger than remaining compressed data')
        compressed_end = min([len(compressed_container), compressed_chunk_start + chunk_size])
        # read after chunk header:
        compressed_current = compressed_chunk_start + 2

        if chunk_flag == 0:
            # MS-OVBA 2.4.1.3.3 Decompressing a RawChunk
            # uncompressed chunk: read the next 4096 bytes as-is
            #TODO: check if there are at least 4096 bytes left
            decompressed_container += compressed_container[compressed_current:compressed_current + 4096]
            compressed_current += 4096
        else:
            # MS-OVBA 2.4.1.3.2 Decompressing a CompressedChunk
            # compressed chunk
            decompressed_chunk_start = len(decompressed_container)
            while compressed_current < compressed_end:
                # MS-OVBA 2.4.1.3.4 Decompressing a TokenSequence
                # logging.debug('compressed_current = %d / compressed_end = %d' % (compressed_current, compressed_end))
                # FlagByte: 8 bits indicating if the following 8 tokens are either literal (1 byte of plain text) or
                # copy tokens (reference to a previous literal token)
                flag_byte = ord(compressed_container[compressed_current])
                compressed_current += 1
                for bit_index in xrange(0, 8):
                    # logging.debug('bit_index=%d / compressed_current=%d / compressed_end=%d' % (bit_index, compressed_current, compressed_end))
                    if compressed_current >= compressed_end:
                        break
                    # MS-OVBA 2.4.1.3.5 Decompressing a Token
                    # MS-OVBA 2.4.1.3.17 Extract FlagBit
                    flag_bit = (flag_byte >> bit_index) & 1
                    #logging.debug('bit_index=%d: flag_bit=%d' % (bit_index, flag_bit))
                    if flag_bit == 0: # LiteralToken
                        # copy one byte directly to output
                        decompressed_container += compressed_container[compressed_current]
                        compressed_current += 1
                    else: # CopyToken
                        # MS-OVBA 2.4.1.3.19.2 Unpack CopyToken
                        copy_token = struct.unpack("<H", compressed_container[compressed_current:compressed_current + 2])[0]
                        #TODO: check this
                        length_mask, offset_mask, bit_count, maximum_length = copytoken_help(
                            len(decompressed_container), decompressed_chunk_start)
                        length = (copy_token & length_mask) + 3
                        temp1 = copy_token & offset_mask
                        temp2 = 16 - bit_count
                        offset = (temp1 >> temp2) + 1
                        #logging.debug('offset=%d length=%d' % (offset, length))
                        copy_source = len(decompressed_container) - offset
                        for index in xrange(copy_source, copy_source + length):
                            decompressed_container += decompressed_container[index]
                        compressed_current += 2
    return decompressed_container


def _extract_vba (ole, vba_root, project_path, dir_path):
    """
    Extract VBA macros from an OleFileIO object.
    Internal function, do not call directly.

    vba_root: path to the VBA root storage, containing the VBA storage and the PROJECT stream
    vba_project: path to the PROJECT stream
    This is a generator, yielding (stream path, VBA filename, VBA source code) for each VBA code stream
    """
    # Open the PROJECT stream:
    project = ole.openstream(project_path)

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

    code_modules = {}

    for line in project:
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
                code_modules[value] = CLASS_EXTENSION
            elif name == 'Module':
                code_modules[value] = MODULE_EXTENSION
            elif name == 'Class':
                code_modules[value] = CLASS_EXTENSION
            elif name == 'BaseClass':
                code_modules[value] = FORM_EXTENSION

    # read data from dir stream (compressed)
    dir_compressed = ole.openstream(dir_path).read()

    def check_value(name, expected, value):
        if expected != value:
            logging.error("invalid value for {0} expected {1:04X} got {2:04X}".format(name, expected, value))

    dir_stream = cStringIO.StringIO(decompress_stream(dir_compressed))

    # PROJECTSYSKIND Record
    PROJECTSYSKIND_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTSYSKIND_Id', 0x0001, PROJECTSYSKIND_Id)
    PROJECTSYSKIND_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTSYSKIND_Size', 0x0004, PROJECTSYSKIND_Size)
    PROJECTSYSKIND_SysKind = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTSYSKIND_SysKind == 0x00:
        logging.debug("16-bit Windows")
    elif PROJECTSYSKIND_SysKind == 0x01:
        logging.debug("32-bit Windows")
    elif PROJECTSYSKIND_SysKind == 0x02:
        logging.debug("Macintosh")
    elif PROJECTSYSKIND_SysKind == 0x03:
        logging.debug("64-bit Windows")
    else:
        logging.error("invalid PROJECTSYSKIND_SysKind {0:04X}".format(PROJECTSYSKIND_SysKind))

    # PROJECTLCID Record
    PROJECTLCID_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTLCID_Id', 0x0002, PROJECTLCID_Id)
    PROJECTLCID_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLCID_Size', 0x0004, PROJECTLCID_Size)
    PROJECTLCID_Lcid = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLCID_Lcid', 0x409, PROJECTLCID_Lcid)

    # PROJECTLCIDINVOKE Record
    PROJECTLCIDINVOKE_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTLCIDINVOKE_Id', 0x0014, PROJECTLCIDINVOKE_Id)
    PROJECTLCIDINVOKE_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLCIDINVOKE_Size', 0x0004, PROJECTLCIDINVOKE_Size)
    PROJECTLCIDINVOKE_LcidInvoke = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLCIDINVOKE_LcidInvoke', 0x409, PROJECTLCIDINVOKE_LcidInvoke)

    # PROJECTCODEPAGE Record
    PROJECTCODEPAGE_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTCODEPAGE_Id', 0x0003, PROJECTCODEPAGE_Id)
    PROJECTCODEPAGE_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTCODEPAGE_Size', 0x0002, PROJECTCODEPAGE_Size)
    PROJECTCODEPAGE_CodePage = struct.unpack("<H", dir_stream.read(2))[0]

    # PROJECTNAME Record
    PROJECTNAME_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTNAME_Id', 0x0004, PROJECTNAME_Id)
    PROJECTNAME_SizeOfProjectName = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTNAME_SizeOfProjectName < 1 or PROJECTNAME_SizeOfProjectName > 128:
        logging.error("PROJECTNAME_SizeOfProjectName value not in range: {0}".format(PROJECTNAME_SizeOfProjectName))
    PROJECTNAME_ProjectName = dir_stream.read(PROJECTNAME_SizeOfProjectName)

    # PROJECTDOCSTRING Record
    PROJECTDOCSTRING_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTDOCSTRING_Id', 0x0005, PROJECTDOCSTRING_Id)
    PROJECTDOCSTRING_SizeOfDocString = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTNAME_SizeOfProjectName > 2000:
        logging.error("PROJECTDOCSTRING_SizeOfDocString value not in range: {0}".format(PROJECTDOCSTRING_SizeOfDocString))
    PROJECTDOCSTRING_DocString = dir_stream.read(PROJECTDOCSTRING_SizeOfDocString)
    PROJECTDOCSTRING_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTDOCSTRING_Reserved', 0x0040, PROJECTDOCSTRING_Reserved)
    PROJECTDOCSTRING_SizeOfDocStringUnicode = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTDOCSTRING_SizeOfDocStringUnicode % 2 != 0:
        logging.error("PROJECTDOCSTRING_SizeOfDocStringUnicode is not even")
    PROJECTDOCSTRING_DocStringUnicode = dir_stream.read(PROJECTDOCSTRING_SizeOfDocStringUnicode)

    # PROJECTHELPFILEPATH Record - MS-OVBA 2.3.4.2.1.7
    PROJECTHELPFILEPATH_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTHELPFILEPATH_Id', 0x0006, PROJECTHELPFILEPATH_Id)
    PROJECTHELPFILEPATH_SizeOfHelpFile1 = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTHELPFILEPATH_SizeOfHelpFile1 > 260:
        logging.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 value not in range: {0}".format(PROJECTHELPFILEPATH_SizeOfHelpFile1))
    PROJECTHELPFILEPATH_HelpFile1 = dir_stream.read(PROJECTHELPFILEPATH_SizeOfHelpFile1)
    PROJECTHELPFILEPATH_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTHELPFILEPATH_Reserved', 0x003D, PROJECTHELPFILEPATH_Reserved)
    PROJECTHELPFILEPATH_SizeOfHelpFile2 = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTHELPFILEPATH_SizeOfHelpFile2 != PROJECTHELPFILEPATH_SizeOfHelpFile1:
        logging.error("PROJECTHELPFILEPATH_SizeOfHelpFile1 does not equal PROJECTHELPFILEPATH_SizeOfHelpFile2")
    PROJECTHELPFILEPATH_HelpFile2 = dir_stream.read(PROJECTHELPFILEPATH_SizeOfHelpFile2)
    if PROJECTHELPFILEPATH_HelpFile2 != PROJECTHELPFILEPATH_HelpFile1:
        logging.error("PROJECTHELPFILEPATH_HelpFile1 does not equal PROJECTHELPFILEPATH_HelpFile2")

    # PROJECTHELPCONTEXT Record
    PROJECTHELPCONTEXT_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTHELPCONTEXT_Id', 0x0007, PROJECTHELPCONTEXT_Id)
    PROJECTHELPCONTEXT_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTHELPCONTEXT_Size', 0x0004, PROJECTHELPCONTEXT_Size)
    PROJECTHELPCONTEXT_HelpContext = struct.unpack("<L", dir_stream.read(4))[0]

    # PROJECTLIBFLAGS Record
    PROJECTLIBFLAGS_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTLIBFLAGS_Id', 0x0008, PROJECTLIBFLAGS_Id)
    PROJECTLIBFLAGS_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLIBFLAGS_Size', 0x0004, PROJECTLIBFLAGS_Size)
    PROJECTLIBFLAGS_ProjectLibFlags = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTLIBFLAGS_ProjectLibFlags', 0x0000, PROJECTLIBFLAGS_ProjectLibFlags)

    # PROJECTVERSION Record
    PROJECTVERSION_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTVERSION_Id', 0x0009, PROJECTVERSION_Id)
    PROJECTVERSION_Reserved = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTVERSION_Reserved', 0x0004, PROJECTVERSION_Reserved)
    PROJECTVERSION_VersionMajor = struct.unpack("<L", dir_stream.read(4))[0]
    PROJECTVERSION_VersionMinor = struct.unpack("<H", dir_stream.read(2))[0]

    # PROJECTCONSTANTS Record
    PROJECTCONSTANTS_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTCONSTANTS_Id', 0x000C, PROJECTCONSTANTS_Id)
    PROJECTCONSTANTS_SizeOfConstants = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTCONSTANTS_SizeOfConstants > 1015:
        logging.error("PROJECTCONSTANTS_SizeOfConstants value not in range: {0}".format(PROJECTCONSTANTS_SizeOfConstants))
    PROJECTCONSTANTS_Constants = dir_stream.read(PROJECTCONSTANTS_SizeOfConstants)
    PROJECTCONSTANTS_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTCONSTANTS_Reserved', 0x003C, PROJECTCONSTANTS_Reserved)
    PROJECTCONSTANTS_SizeOfConstantsUnicode = struct.unpack("<L", dir_stream.read(4))[0]
    if PROJECTCONSTANTS_SizeOfConstantsUnicode % 2 != 0:
        logging.error("PROJECTCONSTANTS_SizeOfConstantsUnicode is not even")
    PROJECTCONSTANTS_ConstantsUnicode = dir_stream.read(PROJECTCONSTANTS_SizeOfConstantsUnicode)

    # array of REFERENCE records
    check = None
    while True:
        check = struct.unpack("<H", dir_stream.read(2))[0]
        logging.debug("reference type = {0:04X}".format(check))
        if check == 0x000F:
            break

        if check == 0x0016:
            # REFERENCENAME
            REFERENCE_Id = check
            REFERENCE_SizeOfName = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCE_Name = dir_stream.read(REFERENCE_SizeOfName)
            REFERENCE_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
            check_value('REFERENCE_Reserved', 0x003E, REFERENCE_Reserved)
            REFERENCE_SizeOfNameUnicode = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCE_NameUnicode = dir_stream.read(REFERENCE_SizeOfNameUnicode)
            continue

        if check == 0x0033:
            # REFERENCEORIGINAL (followed by REFERENCECONTROL)
            REFERENCEORIGINAL_Id = check
            REFERENCEORIGINAL_SizeOfLibidOriginal = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEORIGINAL_LibidOriginal = dir_stream.read(REFERENCEORIGINAL_SizeOfLibidOriginal)
            continue

        if check == 0x002F:
            # REFERENCECONTROL
            REFERENCECONTROL_Id = check
            REFERENCECONTROL_SizeTwiddled = struct.unpack("<L", dir_stream.read(4))[0] # ignore
            REFERENCECONTROL_SizeOfLibidTwiddled = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCECONTROL_LibidTwiddled = dir_stream.read(REFERENCECONTROL_SizeOfLibidTwiddled)
            REFERENCECONTROL_Reserved1 = struct.unpack("<L", dir_stream.read(4))[0] # ignore
            check_value('REFERENCECONTROL_Reserved1', 0x0000, REFERENCECONTROL_Reserved1)
            REFERENCECONTROL_Reserved2 = struct.unpack("<H", dir_stream.read(2))[0] # ignore
            check_value('REFERENCECONTROL_Reserved2', 0x0000, REFERENCECONTROL_Reserved2)
            # optional field
            check2 = struct.unpack("<H", dir_stream.read(2))[0]
            if check2 == 0x0016:
                REFERENCECONTROL_NameRecordExtended_Id = check
                REFERENCECONTROL_NameRecordExtended_SizeofName = struct.unpack("<L", dir_stream.read(4))[0]
                REFERENCECONTROL_NameRecordExtended_Name = dir_stream.read(REFERENCECONTROL_NameRecordExtended_SizeofName)
                REFERENCECONTROL_NameRecordExtended_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
                check_value('REFERENCECONTROL_NameRecordExtended_Reserved', 0x003E, REFERENCECONTROL_NameRecordExtended_Reserved)
                REFERENCECONTROL_NameRecordExtended_SizeOfNameUnicode = struct.unpack("<L", dir_stream.read(4))[0]
                REFERENCECONTROL_NameRecordExtended_NameUnicode = dir_stream.read(REFERENCECONTROL_NameRecordExtended_SizeOfNameUnicode)
                REFERENCECONTROL_Reserved3 = struct.unpack("<H", dir_stream.read(2))[0]
            else:
                REFERENCECONTROL_Reserved3 = check2

            check_value('REFERENCECONTROL_Reserved3', 0x0030, REFERENCECONTROL_Reserved3)
            REFERENCECONTROL_SizeExtended = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCECONTROL_SizeOfLibidExtended = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCECONTROL_LibidExtended = dir_stream.read(REFERENCECONTROL_SizeOfLibidExtended)
            REFERENCECONTROL_Reserved4 = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCECONTROL_Reserved5 = struct.unpack("<H", dir_stream.read(2))[0]
            REFERENCECONTROL_OriginalTypeLib = dir_stream.read(16)
            REFERENCECONTROL_Cookie = struct.unpack("<L", dir_stream.read(4))[0]
            continue

        if check == 0x000D:
            # REFERENCEREGISTERED
            REFERENCEREGISTERED_Id = check
            REFERENCEREGISTERED_Size = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEREGISTERED_SizeOfLibid = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEREGISTERED_Libid = dir_stream.read(REFERENCEREGISTERED_SizeOfLibid)
            REFERENCEREGISTERED_Reserved1 = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('REFERENCEREGISTERED_Reserved1', 0x0000, REFERENCEREGISTERED_Reserved1)
            REFERENCEREGISTERED_Reserved2 = struct.unpack("<H", dir_stream.read(2))[0]
            check_value('REFERENCEREGISTERED_Reserved2', 0x0000, REFERENCEREGISTERED_Reserved2)
            continue

        if check == 0x000E:
            # REFERENCEPROJECT
            REFERENCEPROJECT_Id = check
            REFERENCEPROJECT_Size = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEPROJECT_SizeOfLibidAbsolute = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEPROJECT_LibidAbsolute = dir_stream.read(REFERENCEPROJECT_SizeOfLibidAbsolute)
            REFERENCEPROJECT_SizeOfLibidRelative = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEPROJECT_LibidRelative = dir_stream.read(REFERENCEPROJECT_SizeOfLibidRelative)
            REFERENCEPROJECT_MajorVersion = struct.unpack("<L", dir_stream.read(4))[0]
            REFERENCEPROJECT_MinorVersion = struct.unpack("<H", dir_stream.read(2))[0]
            continue

        logging.error('invalid or unknown check Id {0:04X}'.format(check))
        sys.exit(0)

    PROJECTMODULES_Id = check #struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTMODULES_Id', 0x000F, PROJECTMODULES_Id)
    PROJECTMODULES_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTMODULES_Size', 0x0002, PROJECTMODULES_Size)
    PROJECTMODULES_Count = struct.unpack("<H", dir_stream.read(2))[0]
    PROJECTMODULES_ProjectCookieRecord_Id = struct.unpack("<H", dir_stream.read(2))[0]
    check_value('PROJECTMODULES_ProjectCookieRecord_Id', 0x0013, PROJECTMODULES_ProjectCookieRecord_Id)
    PROJECTMODULES_ProjectCookieRecord_Size = struct.unpack("<L", dir_stream.read(4))[0]
    check_value('PROJECTMODULES_ProjectCookieRecord_Size', 0x0002, PROJECTMODULES_ProjectCookieRecord_Size)
    PROJECTMODULES_ProjectCookieRecord_Cookie = struct.unpack("<H", dir_stream.read(2))[0]

    logging.debug("parsing {0} modules".format(PROJECTMODULES_Count))
    for x in xrange(0, PROJECTMODULES_Count):
        MODULENAME_Id = struct.unpack("<H", dir_stream.read(2))[0]
        check_value('MODULENAME_Id', 0x0019, MODULENAME_Id)
        MODULENAME_SizeOfModuleName = struct.unpack("<L", dir_stream.read(4))[0]
        MODULENAME_ModuleName = dir_stream.read(MODULENAME_SizeOfModuleName)
        # account for optional sections
        section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x0047:
            MODULENAMEUNICODE_Id = section_id
            MODULENAMEUNICODE_SizeOfModuleNameUnicode = struct.unpack("<L", dir_stream.read(4))[0]
            MODULENAMEUNICODE_ModuleNameUnicode = dir_stream.read(MODULENAMEUNICODE_SizeOfModuleNameUnicode)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x001A:
            MODULESTREAMNAME_id = section_id
            MODULESTREAMNAME_SizeOfStreamName = struct.unpack("<L", dir_stream.read(4))[0]
            MODULESTREAMNAME_StreamName = dir_stream.read(MODULESTREAMNAME_SizeOfStreamName)
            MODULESTREAMNAME_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
            check_value('MODULESTREAMNAME_Reserved', 0x0032, MODULESTREAMNAME_Reserved)
            MODULESTREAMNAME_SizeOfStreamNameUnicode = struct.unpack("<L", dir_stream.read(4))[0]
            MODULESTREAMNAME_StreamNameUnicode = dir_stream.read(MODULESTREAMNAME_SizeOfStreamNameUnicode)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x001C:
            MODULEDOCSTRING_Id = section_id
            check_value('MODULEDOCSTRING_Id', 0x001C, MODULEDOCSTRING_Id)
            MODULEDOCSTRING_SizeOfDocString = struct.unpack("<L", dir_stream.read(4))[0]
            MODULEDOCSTRING_DocString = dir_stream.read(MODULEDOCSTRING_SizeOfDocString)
            MODULEDOCSTRING_Reserved = struct.unpack("<H", dir_stream.read(2))[0]
            check_value('MODULEDOCSTRING_Reserved', 0x0048, MODULEDOCSTRING_Reserved)
            MODULEDOCSTRING_SizeOfDocStringUnicode = struct.unpack("<L", dir_stream.read(4))[0]
            MODULEDOCSTRING_DocStringUnicode = dir_stream.read(MODULEDOCSTRING_SizeOfDocStringUnicode)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x0031:
            MODULEOFFSET_Id = section_id
            check_value('MODULEOFFSET_Id', 0x0031, MODULEOFFSET_Id)
            MODULEOFFSET_Size = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULEOFFSET_Size', 0x0004, MODULEOFFSET_Size)
            MODULEOFFSET_TextOffset = struct.unpack("<L", dir_stream.read(4))[0]
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x001E:
            MODULEHELPCONTEXT_Id = section_id
            check_value('MODULEHELPCONTEXT_Id', 0x001E, MODULEHELPCONTEXT_Id)
            MODULEHELPCONTEXT_Size = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULEHELPCONTEXT_Size', 0x0004, MODULEHELPCONTEXT_Size)
            MODULEHELPCONTEXT_HelpContext = struct.unpack("<L", dir_stream.read(4))[0]
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x002C:
            MODULECOOKIE_Id = section_id
            check_value('MODULECOOKIE_Id', 0x002C, MODULECOOKIE_Id)
            MODULECOOKIE_Size = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULECOOKIE_Size', 0x0002, MODULECOOKIE_Size)
            MODULECOOKIE_Cookie = struct.unpack("<H", dir_stream.read(2))[0]
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x0021 or section_id == 0x0022:
            MODULETYPE_Id = section_id
            MODULETYPE_Reserved = struct.unpack("<L", dir_stream.read(4))[0]
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x0025:
            MODULEREADONLY_Id = section_id
            check_value('MODULEREADONLY_Id', 0x0025, MODULEREADONLY_Id)
            MODULEREADONLY_Reserved = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULEREADONLY_Reserved', 0x0000, MODULEREADONLY_Reserved)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x0028:
            MODULEPRIVATE_Id = section_id
            check_value('MODULEPRIVATE_Id', 0x0028, MODULEPRIVATE_Id)
            MODULEPRIVATE_Reserved = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULEPRIVATE_Reserved', 0x0000, MODULEPRIVATE_Reserved)
            section_id = struct.unpack("<H", dir_stream.read(2))[0]
        if section_id == 0x002B: # TERMINATOR
            MODULE_Reserved = struct.unpack("<L", dir_stream.read(4))[0]
            check_value('MODULE_Reserved', 0x0000, MODULE_Reserved)
            section_id = None
        if section_id != None:
            logging.warning('unknown or invalid module section id {0:04X}'.format(section_id))

        logging.debug('Project CodePage = %d' % PROJECTCODEPAGE_CodePage)
        vba_codec = 'cp%d' % PROJECTCODEPAGE_CodePage
        logging.debug("ModuleName = {0}".format(MODULENAME_ModuleName))
        logging.debug("StreamName = {0}".format(repr(MODULESTREAMNAME_StreamName)))
        streamname_unicode = MODULESTREAMNAME_StreamName.decode(vba_codec)
        logging.debug("StreamName.decode('%s') = %s" % (vba_codec, repr(streamname_unicode)))
        logging.debug("StreamNameUnicode = {0}".format(repr(MODULESTREAMNAME_StreamNameUnicode)))
        logging.debug("TextOffset = {0}".format(MODULEOFFSET_TextOffset))

        code_path = vba_root + u'VBA/' + streamname_unicode
        #TODO: test if stream exists
        logging.debug('opening VBA code stream %s' % repr(code_path))
        code_data = ole.openstream(code_path).read()
        logging.debug("length of code_data = {0}".format(len(code_data)))
        logging.debug("offset of code_data = {0}".format(MODULEOFFSET_TextOffset))
        code_data = code_data[MODULEOFFSET_TextOffset:]
        if len(code_data) > 0:
            code_data = decompress_stream(code_data)
            # case-insensitive search in the code_modules dict to find the file extension:
            filext = code_modules.get(MODULENAME_ModuleName.lower(), 'bin')
            filename = '{0}.{1}'.format(MODULENAME_ModuleName, filext)
            #TODO: also yield the codepage so that callers can decode it properly
            yield (code_path, filename, code_data)
            # print '-'*79
            # print filename
            # print ''
            # print code_data
            # print ''
            logging.debug('extracted file {0}'.format(filename))
        else:
            logging.warning("module stream {0} has code data length 0".format(MODULESTREAMNAME_StreamName))
    return


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
    for description, keywords in AUTOEXEC_KEYWORDS.items():
        for keyword in keywords:
            #TODO: if keyword is already a compiled regex, use it as-is
            # search using regex to detect word boundaries:
            if re.search(r'(?i)\b'+keyword+r'\b', vba_code):
            #if keyword.lower() in vba_code:
                results.append((keyword, description+obf_text))
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
            if re.search(r'(?i)\b'+keyword+r'\b', vba_code):
            #if keyword.lower() in vba_code:
                results.append((keyword, description+obf_text))
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
            if value not in found:
                results.append((pattern_type+obf_text, value))
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
            decoded = binascii.unhexlify(value)
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
                decoded = base64.b64decode(value)
                results.append((value, decoded))
                found.add(value)
            except:
                # if an exception occurs, it is likely not a base64-encoded string
                pass
    return results


def detect_dridex_strings(vba_code):
    """
    Detect if the VBA code contains strings obfuscated with a specific algorithm found in Dridex samples.

    :param vba_code: str, VBA source code
    :return: list of str tuples (encoded string, decoded string)
    """
    from thirdparty.DridexUrlDecoder.DridexUrlDecoder import DridexUrlDecode
    results = []
    found = set()
    for match in re_dridex_string.finditer(vba_code):
        value = match.group()[1:-1]
        # check it is not just a hex string:
        if not re_nothex_check.search(value):
            continue
        if value not in found:
            try:
                decoded = DridexUrlDecode(value)
                results.append((value, decoded))
                found.add(value)
            except:
                # if an exception occurs, it is likely not a dridex-encoded string
                pass
    return results


class VBA_Scanner (object):
    """
    Class to scan the source code of a VBA module to find obfuscated strings,
    suspicious keywords, IOCs, auto-executable macros, etc.
    """

    def __init__(self, vba_code):
        """
        VBA_Scanner constructor

        :param vba_code: str, VBA source code to be analyzed
        """
        self.code = vba_code
        self.code_hex = ''
        self.code_hex_rev = ''
        self.code_rev_hex = ''
        self.code_base64 = ''
        self.code_dridex = ''


    def scan(self, include_decoded_strings=False):
        """
        Analyze the provided VBA code to detect suspicious keywords,
        auto-executable macros, IOC patterns, obfuscation patterns
        such as hex-encoded strings.

        :param include_decoded_strings: bool, if True, all encoded strings will be included with their decoded content.
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
            self.code_hex += '\n'+decoded
            # if the code contains "StrReverse", also append the hex strings in reverse order:
            if self.strReverse:
                # StrReverse after hex decoding:
                self.code_hex_rev += '\n'+decoded[::-1]
                # StrReverse before hex decoding:
                self.code_rev_hex += '\n'+binascii.unhexlify(encoded[::-1])
                #example: https://malwr.com/analysis/NmFlMGI4YTY1YzYyNDkwNTg1ZTBiZmY5OGI3YjlhYzU/
        #TODO: also append the full code reversed if StrReverse? (risk of false positives?)
        # Detect Base64-encoded strings
        self.base64_strings = detect_base64_strings(self.code)
        for encoded, decoded in self.base64_strings:
            self.code_base64 += '\n'+decoded
        # Detect Dridex-encoded strings
        self.dridex_strings = detect_dridex_strings(self.code)
        for encoded, decoded in self.dridex_strings:
            self.code_dridex += '\n'+decoded
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
        for keyword, description in self.autoexec_keywords:
            results.append(('AutoExec', keyword, description))
        for keyword, description in self.suspicious_keywords:
            results.append(('Suspicious', keyword, description))
        for pattern_type, value in self.iocs:
            results.append(('IOC', value, pattern_type))
        if include_decoded_strings:
            for encoded, decoded in self.hex_strings:
                results.append(('Hex String', repr(decoded), encoded))
            for encoded, decoded in self.base64_strings:
                results.append(('Base64 String', repr(decoded), encoded))
            for encoded, decoded in self.dridex_strings:
                results.append(('Dridex string', repr(decoded), encoded))
        return results

    def scan_summary(self):
        """
        Analyze the provided VBA code to detect suspicious keywords,
        auto-executable macros, IOC patterns, obfuscation patterns
        such as hex-encoded strings.

        :return: tuple with the number of items found for each category:
            (autoexec, suspicious, IOCs, hex, base64, dridex)
        """
        self.scan()
        return (len(self.autoexec_keywords), len(self.suspicious_keywords),
            len(self.iocs), len(self.hex_strings), len(self.base64_strings),
            len(self.dridex_strings))



def scan_vba(vba_code, include_decoded_strings):
    """
    Analyze the provided VBA code to detect suspicious keywords,
    auto-executable macros, IOC patterns, obfuscation patterns
    such as hex-encoded strings.
    (shortcut for VBA_Scanner(vba_code).scan())

    :param vba_code: str, VBA source code to be analyzed
    :param include_decoded_strings: bool, if True all encoded strings will be included with their decoded content.
    :return: list of tuples (type, keyword, description)
    (type = 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String' or 'Dridex String')
    """
    return VBA_Scanner(vba_code).scan(include_decoded_strings)


#=== CLASSES =================================================================

class VBA_Parser(object):
    """
    Class to parse MS Office files, to detect VBA macros and extract VBA source code
    Supported file formats:
    - Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
    - Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
    - PowerPoint 2007+ (.pptm, .ppsm)
    """

    def __init__(self, filename, data=None):
        """
        Constructor for VBA_Parser

        :param filename: filename or path of file to parse, or file-like object

        :param data: None or bytes str, if None the file will be read from disk (or from the file-like object).
        If data is provided as a bytes string, it will be parsed as the content of the file in memory,
        and not read from disk. Note: files must be read in binary mode, i.e. open(f, 'rb').
        """
        #TODO: filename should only be a string, data should be used for the file-like object
        #TODO: filename should be mandatory, optional data is a string or file-like object
        #TODO: also support olefile and zipfile as input
        if data is None:
            # open file from disk:
            _file = filename
        else:
            # file already read in memory, make it a file-like object for zipfile:
            _file = cStringIO.StringIO(data)
        #self.file = _file
        self.ole_file = None
        self.ole_subfiles = []
        self.filename = filename
        self.type = None
        self.vba_projects = None
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
            logging.info('Parsing OLE file %s' % self.filename)
            # Open and parse the OLE file, using unicode for path names:
            self.ole_file = olefile.OleFileIO(_file, path_encoding=None)
            self.type = TYPE_OLE
            #TODO: raise TypeError if this is a Powerpoint 97 file, since VBA macros cannot be detected yet
        elif zipfile.is_zipfile(_file):
            # This looks like a zip file, need to look for vbaProject.bin inside
            # It can be any OLE file inside the archive
            #...because vbaProject.bin can be renamed:
            # see http://www.decalage.info/files/JCV07_Lagadec_OpenDocument_OpenXML_v4_decalage.pdf#page=18
            logging.info('Opening ZIP/OpenXML file %s' % self.filename)
            self.type = TYPE_OpenXML
            z = zipfile.ZipFile(_file)
            #TODO: check if this is actually an OpenXML file
            #TODO: if the zip file is encrypted, suggest to use the -z option, or try '-z infected' automatically?
            # check each file within the zip if it is an OLE file, by reading its magic:
            for subfile in z.namelist():
                magic = z.open(subfile).read(len(olefile.MAGIC))
                if magic == olefile.MAGIC:
                    logging.debug('Opening OLE file %s within zip' % subfile)
                    ole_data = z.open(subfile).read()
                    try:
                        self.ole_subfiles.append(VBA_Parser(filename=subfile, data=ole_data))
                    except:
                        logging.debug('%s is not a valid OLE file' % subfile)
                        continue
            z.close()
        else:
            msg = '%s is not an OLE nor an OpenXML file, cannot extract VBA Macros.' % self.filename
            logging.error(msg)
            raise TypeError(msg)

    def find_vba_projects (self):
        """
        Finds all the VBA projects stored in an OLE file.

        Return None if the file is not OLE but OpenXML.
        Return a list of tuples (vba_root, project_path, dir_path) for each VBA project.
        vba_root is the path of the root OLE storage containing the VBA project,
        including a trailing slash unless it is the root of the OLE file.
        project_path is the path of the OLE stream named "PROJECT" within the VBA project.
        dir_path is the path of the OLE stream named "VBA/dir" within the VBA project.

        If this function returns an empty list for one of the supported formats
        (i.e. Word, Excel, Powerpoint except Powerpoint 97-2003), then the
        file does not contain VBA macros.

        :return: None if OpenXML file, list of tuples (vba_root, project_path, dir_path)
        for each VBA project found if OLE file
        """
        # if the file is not OLE but OpenXML, return None:
        if self.ole_file is None:
            return None

        # if this method has already been called, return previous result:
        if self.vba_projects is not None:
            return self.vba_projects

        # Find the VBA project root (different in MS Word, Excel, etc):
        # - Word 97-2003: Macros
        # - Excel 97-2003: _VBA_PROJECT_CUR
        # - PowerPoint 97-2003: not supported yet (different file structure)
        # - Word 2007+: word/vbaProject.bin in zip archive, then the VBA project is the root of vbaProject.bin.
        # - Excel 2007+: xl/vbaProject.bin in zip archive, then same as Word
        # - PowerPoint 2007+: ppt/vbaProject.bin in zip archive, then same as Word
        # - Visio 2007: not supported yet (different file structure)

        # According to MS-OVBA section 2.2.1:
        # - the VBA project root storage MUST contain a VBA storage and a PROJECT stream
        # - The root/VBA storage MUST contain a _VBA_PROJECT stream and a dir stream
        # - all names are case-insensitive

        # start with an empty list:
        self.vba_projects = []
        # Look for any storage containing those storage/streams:
        ole = self.ole_file
        for storage in ole.listdir(streams=False, storages=True):
            # Look for a storage ending with "VBA":
            if storage[-1].upper() == 'VBA':
                logging.debug('Found VBA storage: %s' % ('/'.join(storage)))
                vba_root = '/'.join(storage[:-1])
                # Add a trailing slash to vba_root, unless it is the root of the OLE file:
                # (used later to append all the child streams/storages)
                if vba_root != '':
                    vba_root += '/'
                logging.debug('Checking vba_root="%s"' % vba_root)

                def check_vba_stream(ole, vba_root, stream_path):
                    full_path = vba_root + stream_path
                    if ole.exists(full_path) and ole.get_type(full_path) == olefile.STGTY_STREAM:
                        logging.debug('Found %s stream: %s' % (stream_path, full_path))
                        return full_path
                    else:
                        logging.debug('Missing %s stream, this is not a valid VBA project structure' % stream_path)
                        return False

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
                logging.debug('VBA root storage: "%s"' % vba_root)
                # append the results to the list as a tuple for later use:
                self.vba_projects.append((vba_root, project_path, dir_path))
        return self.vba_projects

    def detect_vba_macros(self):
        """
        Detect the potential presence of VBA macros in the file, by checking
        if it contains VBA projects. Both OLE and OpenXML files are supported.

        Important: for now, results are accurate only for Word, Excel and PowerPoint
        EXCEPT Powerpoint 97-2003, which has a different structure for VBA.

        Note: this method does NOT attempt to check the actual presence or validity
        of VBA macro source code, so there might be false positives.
        It may also detect VBA macros in files embedded within the main file,
        for example an Excel workbook with macros embedded into a Word
        document without macros may be detected, without distinction.

        :return: bool, True if at least one VBA project has been found, False otherwise
        """
        #TODO: return None or raise exception if format not supported like PPT 97-2003
        #TODO: return the number of VBA projects found instead of True/False?
        # if OpenXML, check all the OLE subfiles:
        if self.ole_file is None:
            for ole_subfile in self.ole_subfiles:
                if ole_subfile.detect_vba_macros():
                    return True
            return False
        # otherwise it's an OLE file, find VBA projects:
        vba_projects = self.find_vba_projects()
        if len(vba_projects) == 0:
            return False
        else:
            return True


    def extract_macros (self):
        """
        Extract and decompress source code for each VBA macro found in the file

        Iterator: yields (filename, stream_path, vba_filename, vba_code) for each VBA macro found
        If the file is OLE, filename is the path of the file.
        If the file is OpenXML, filename is the path of the OLE subfile containing VBA macros
        within the zip archive, e.g. word/vbaProject.bin.
        """
        if self.ole_file is None:
            for ole_subfile in self.ole_subfiles:
                for results in ole_subfile.extract_macros():
                    yield results
        else:
            self.find_vba_projects()
            for vba_root, project_path, dir_path in self.vba_projects:
                # extract all VBA macros from that VBA root storage:
                for stream_path, vba_filename, vba_code in _extract_vba(self.ole_file, vba_root, project_path, dir_path):
                    yield (self.filename, stream_path, vba_filename, vba_code)


    def close(self):
        """
        Close all the open files. This method must be called after usage, if
        the application is opening many files.
        """
        if self.ole_file is None:
            for ole_subfile in self.ole_subfiles:
                ole_subfile.close()
        else:
            self.ole_file.close()


def print_analysis(vba_code, show_decoded_strings=False):
    """
    Analyze the provided VBA code, and print the results in a table

    :param vba_code: str, VBA source code to be analyzed
    :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
    :return: None
    """
    results = scan_vba(vba_code, show_decoded_strings)
    if results:
        t = prettytable.PrettyTable(('Type', 'Keyword', 'Description'))
        t.align = 'l'
        t.max_width['Type'] = 10
        t.max_width['Keyword'] = 20
        t.max_width['Description'] = 39
        for kw_type, keyword, description in results:
            t.add_row((kw_type, keyword, description))
        print t
    else:
        print 'No suspicious keyword or IOC found.'



def process_file (container, filename, data, show_decoded_strings=False):
    """
    Process a single file

    :param container: str, path and filename of container if the file is within
    a zip archive, None otherwise.
    :param filename: str, path and filename of file on disk, or within the container.
    :param data: bytes, content of the file if it is in a container, None if it is a file on disk.
    :param show_decoded_strings: bool, if True hex-encoded strings will be displayed with their decoded content.
    """
    #TODO: replace print by writing to a provided output file (sys.stdout by default)
    if container:
        display_filename = '%s in %s' % (filename, container)
    else:
        display_filename = filename
    print '='*79
    print 'FILE:', display_filename
    try:
        #TODO: handle olefile errors, when an OLE file is malformed
        vba = VBA_Parser(filename, data)
        print 'Type:', vba.type
        if vba.detect_vba_macros():
            #print 'Contains VBA Macros:'
            for (subfilename, stream_path, vba_filename, vba_code) in vba.extract_macros():
                # hide attribute lines:
                #TODO: option to disable attribute filtering
                vba_code_filtered = filter_vba(vba_code)
                print '-'*79
                print 'VBA MACRO %s ' % vba_filename
                print 'in file: %s - OLE stream: %s' % (subfilename, repr(stream_path))
                print '- '*39
                # detect empty macros:
                if vba_code_filtered.strip() == '':
                    print '(empty macro)'
                else:
                    print vba_code_filtered
                    print '- '*39
                    print 'ANALYSIS:'
                    # analyse the whole code, filtered to avoid false positives:
                    print_analysis(vba_code_filtered, show_decoded_strings)
        else:
            print 'No VBA macros found.'
    except: #TypeError:
        #raise
        #TODO: print more info if debug mode
        #print sys.exc_value
        # display the exception with full stack trace for debugging, but do not stop:
        traceback.print_exc()
    print ''


def process_file_triage (container, filename, data):
    """
    Process a single file

    :param container: str, path and filename of container if the file is within
    a zip archive, None otherwise.
    :param filename: str, path and filename of file on disk, or within the container.
    :param data: bytes, content of the file if it is in a container, None if it is a file on disk.
    """
    #TODO: replace print by writing to a provided output file (sys.stdout by default)
    nb_macros = 0
    nb_autoexec = 0
    nb_suspicious = 0
    nb_iocs = 0
    nb_hexstrings = 0
    nb_base64strings = 0
    nb_dridexstrings = 0
    # ftype = 'Other'
    message = ''
    try:
        #TODO: handle olefile errors, when an OLE file is malformed
        vba = VBA_Parser(filename, data)
        if vba.detect_vba_macros():
            for (subfilename, stream_path, vba_filename, vba_code) in vba.extract_macros():
                nb_macros += 1
                if vba_code.strip() != '':
                    # analyse the whole code, filtered to avoid false positives:
                    scanner = VBA_Scanner(filter_vba(vba_code))
                    autoexec, suspicious, iocs, hexstrings, base64strings, dridex = scanner.scan_summary()
                    nb_autoexec += autoexec
                    nb_suspicious += suspicious
                    nb_iocs += iocs
                    nb_hexstrings += hexstrings
                    nb_base64strings += base64strings
                    nb_dridexstrings += dridex
        if vba.type == TYPE_OLE:
            flags = 'OLE:'
        else:
            flags = 'OpX:'
        macros = autoexec = suspicious = iocs = hexstrings = base64obf = dridex = '-'
        if nb_macros: macros = 'M'
        if nb_autoexec: autoexec = 'A'
        if nb_suspicious: suspicious = 'S'
        if nb_iocs: iocs = 'I'
        if nb_hexstrings: hexstrings = 'H'
        if nb_base64strings: base64obf = 'B'
        if nb_dridexstrings: dridex = 'D'
        flags += '%s%s%s%s%s%s%s' % (macros, autoexec, suspicious, iocs, hexstrings,
            base64obf, dridex)

        # macros = autoexec = suspicious = iocs = hexstrings = 'no'
        # if nb_macros: macros = 'YES:%d' % nb_macros
        # if nb_autoexec: autoexec = 'YES:%d' % nb_autoexec
        # if nb_suspicious: suspicious = 'YES:%d' % nb_suspicious
        # if nb_iocs: iocs = 'YES:%d' % nb_iocs
        # if nb_hexstrings: hexstrings = 'YES:%d' % nb_hexstrings
        # # 2nd line = info
        # print '%-8s %-7s %-7s %-7s %-7s %-7s' % (vba.type, macros, autoexec, suspicious, iocs, hexstrings)
    except TypeError:
        # file type not OLE nor OpenXML
        flags = '?'
        message = 'File format not supported'
    except:
        # another error occurred
        #raise
        #TODO: print more info if debug mode
        #TODO: distinguish real errors from incorrect file types
        flags = '!ERROR'
        message = sys.exc_value
    line = '%-11s %s' % (flags, filename)
    if message:
        line += ' - %s' % message
    print line

    # t = prettytable.PrettyTable(('filename', 'type', 'macros', 'autoexec', 'suspicious', 'ioc', 'hexstrings'),
    #     header=False, border=False)
    # t.align = 'l'
    # t.max_width['filename'] = 30
    # t.max_width['type'] = 10
    # t.max_width['macros'] = 6
    # t.max_width['autoexec'] = 6
    # t.max_width['suspicious'] = 6
    # t.max_width['ioc'] = 6
    # t.max_width['hexstrings'] = 6
    # t.add_row((filename, ftype, macros, autoexec, suspicious, iocs, hexstrings))
    # print t

def main_triage_quick():
    pass

#=== MAIN =====================================================================

def main():
    """
    Main function, called when olevba is run from the command line
    """
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
    parser.add_option("-t", action="store_true", dest="triage_mode",
        help='triage mode, display results as a summary table (default for multiple files)')
    parser.add_option("-d", action="store_true", dest="detailed_mode",
        help='detailed mode, display full results (default for single file)')
    parser.add_option("-i", "--input", dest='input', type='str', default=None,
        help='input file containing VBA source code to be analyzed (no parsing)')
    parser.add_option("--decode", action="store_true", dest="show_decoded_strings",
        help='display all the obfuscated strings with their decoded content (Hex, Base64, StrReverse, Dridex).')

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0 and not options.input:
        print __doc__
        parser.print_help()
        sys.exit()

    logging.basicConfig(format='%(levelname)s: %(message)s', level=logging.WARNING) #INFO)
    # For now, all logging is disabled:
    logging.disable(logging.CRITICAL)

    if options.input:
        # input file provided with VBA source code to be analyzed directly:
        print 'Analysis of VBA source code from %s:' % options.input
        vba_code = open(options.input).read()
        print_analysis(vba_code, show_decoded_strings=options.show_decoded_strings)
        sys.exit()

    # print '%-8s %-7s %-7s %-7s %-7s %-7s' % ('Type', 'Macros', 'AutoEx', 'Susp.', 'IOCs', 'HexStr')
    # print '%-8s %-7s %-7s %-7s %-7s %-7s' % ('-'*8, '-'*7, '-'*7, '-'*7, '-'*7, '-'*7)
    if not options.detailed_mode or options.triage_mode:
        print '%-11s %-65s' % ('Flags', 'Filename')
        print '%-11s %-65s' % ('-'*11, '-'*65)
    previous_container = None
    count = 0
    container = filename = data = None
    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
        zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        if options.detailed_mode and not options.triage_mode:
            # fully detailed output
            process_file(container, filename, data, show_decoded_strings=options.show_decoded_strings)
        else:
            # print container name when it changes:
            if container != previous_container:
                if container is not None:
                    print '\nFiles in %s:' % container
                previous_container = container
            # summarized output for triage:
            process_file_triage(container, filename, data)
        count += 1
    if not options.detailed_mode or options.triage_mode:
        print '\n(Flags: OpX=OpenXML, M=Macros, A=Auto-executable, S=Suspicious keywords, I=IOCs, H=Hex strings, B=Base64 strings, D=Dridex strings, ?=Unknown)\n'

    if count == 1 and not options.triage_mode and not options.detailed_mode:
        # if options -t and -d were not specified and it's a single file, print details:
        #TODO: avoid doing the analysis twice by storing results
        process_file(container, filename, data, show_decoded_strings=options.show_decoded_strings)

if __name__ == '__main__':
    main()

# This was coded while listening to "Dust" from I Love You But I've Chosen Darkness