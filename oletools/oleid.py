#!/usr/bin/env python
"""
oleid.py

oleid is a script to analyze OLE files such as MS Office documents (e.g. Word,
Excel), to detect specific characteristics that could potentially indicate that
the file is suspicious or malicious, in terms of security (e.g. malware).
For example it can detect VBA macros, embedded Flash objects, fragmentation.
The results can be displayed or returned as XML for further processing.

Usage: oleid.py <file>

oleid project website: http://www.decalage.info/python/oleid

oleid is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# oleid is copyright (c) 2012-2015, Philippe Lagadec (http://www.decalage.info)
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
# 2012-10-29 v0.01 PL: - first version
# 2014-11-29 v0.02 PL: - use olefile instead of OleFileIO_PL
#                      - improved usage display with -h
# 2014-11-30 v0.03 PL: - improved output with prettytable

__version__ = '0.03'


#------------------------------------------------------------------------------
# TODO:
# + extract relevant metadata: codepage, author, application, timestamps, etc
# - detect RTF and OpenXML
# - fragmentation
# - OLE package
# - entropy
# - detect PE header?
# - detect NOPs?
# - list type of each object in object pool?
# - criticality for each indicator?: info, low, medium, high
# - support wildcards with glob?
# - verbose option
# - csv, xml output


#=== IMPORTS =================================================================

import optparse, sys, os, re, zlib, struct
import thirdparty.olefile as olefile
from thirdparty.prettytable import prettytable


#=== FUNCTIONS ===============================================================

def detect_flash (data):
    """
    Detect Flash objects (SWF files) within a binary string of data
    return a list of (start_index, length, compressed) tuples, or [] if nothing
    found.

    Code inspired from xxxswf.py by Alexander Hanel (but significantly reworked)
    http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html
    """
    #TODO: report
    found = []
    for match in re.finditer('CWS|FWS', data):
        start = match.start()
        if start+8 > len(data):
            # header size larger than remaining data, this is not a SWF
            continue
        #TODO: one struct.unpack should be simpler
        # Read Header
        header = data[start:start+3]
        # Read Version
        ver = struct.unpack('<b', data[start+3])[0]
        # Error check for version above 20
        #TODO: is this accurate? (check SWF specifications)
        if ver > 20:
            continue
        # Read SWF Size
        size = struct.unpack('<i', data[start+4:start+8])[0]
        if start+size > len(data) or size < 1024:
            # declared size larger than remaining data, this is not a SWF
            # or declared size too small for a usual SWF
            continue
        # Read SWF into buffer. If compressed read uncompressed size.
        swf = data[start:start+size]
        compressed = False
        if 'CWS' in header:
            compressed = True
            # compressed SWF: data after header (8 bytes) until the end is
            # compressed with zlib. Attempt to decompress it to check if it is
            # valid
            compressed_data = swf[8:]
            try:
                zlib.decompress(compressed_data)
            except:
                continue
        # else we don't check anything at this stage, we only assume it is a
        # valid SWF. So there might be false positives for uncompressed SWF.
        found.append((start, size, compressed))
        #print 'Found SWF start=%x, length=%d' % (start, size)
    return found


#=== CLASSES =================================================================

class Indicator (object):

    def __init__(self, _id, value=None, _type=bool, name=None, description=None):
        self.id = _id
        self.value = value
        self.type = _type
        self.name = name
        if name == None:
            self.name = _id
        self.description = description


class OleID:

    def __init__(self, filename):
        self.filename = filename
        self.indicators = []

    def check(self):
        # check if it is actually an OLE file:
        oleformat = Indicator('ole_format', True, name='OLE format')
        self.indicators.append(oleformat)
        if not olefile.isOleFile(self.filename):
            oleformat.value = False
            return self.indicators
        # parse file:
        self.ole = olefile.OleFileIO(self.filename)
        # checks:
        self.check_properties()
        self.check_encrypted()
        self.check_word()
        self.check_excel()
        self.check_powerpoint()
        self.check_visio()
        self.check_ObjectPool()
        self.check_flash()
        self.ole.close()
        return self.indicators

    def check_properties (self):
        suminfo = Indicator('has_suminfo', False, name='Has SummaryInformation stream')
        self.indicators.append(suminfo)
        appname = Indicator('appname', 'unknown', _type=str, name='Application name')
        self.indicators.append(appname)
        self.suminfo = {}
        # check stream SummaryInformation
        if self.ole.exists("\x05SummaryInformation"):
            suminfo.value = True
            self.suminfo = self.ole.getproperties("\x05SummaryInformation")
            # check application name:
            appname.value = self.suminfo.get(0x12, 'unknown')

    def check_encrypted (self):
        # we keep the pointer to the indicator, can be modified by other checks:
        self.encrypted = Indicator('encrypted', False, name='Encrypted')
        self.indicators.append(self.encrypted)
        # check if bit 1 of security field = 1:
        # (this field may be missing for Powerpoint2000, for example)
        if 0x13 in self.suminfo:
            if self.suminfo[0x13] & 1:
                self.encrypted.value = True

    def check_word (self):
        word = Indicator('word', False, name='Word Document',
            description='Contains a WordDocument stream, very likely to be a Microsoft Word Document.')
        self.indicators.append(word)
        self.macros = Indicator('vba_macros', False, name='VBA Macros')
        self.indicators.append(self.macros)
        if self.ole.exists('WordDocument'):
            word.value = True
            # check for Word-specific encryption flag:
            s = self.ole.openstream(["WordDocument"])
            # pass header 10 bytes
            s.read(10)
            # read flag structure:
            temp16 = struct.unpack("H", s.read(2))[0]
            fEncrypted = (temp16 & 0x0100) >> 8
            if fEncrypted:
                self.encrypted.value = True
            s.close()
            # check for VBA macros:
            if self.ole.exists('Macros'):
                self.macros.value = True

    def check_excel (self):
        excel = Indicator('excel', False, name='Excel Workbook',
            description='Contains a Workbook or Book stream, very likely to be a Microsoft Excel Workbook.')
        self.indicators.append(excel)
        #self.macros = Indicator('vba_macros', False, name='VBA Macros')
        #self.indicators.append(self.macros)
        if self.ole.exists('Workbook') or self.ole.exists('Book'):
            excel.value = True
            # check for VBA macros:
            if self.ole.exists('_VBA_PROJECT_CUR'):
                self.macros.value = True

    def check_powerpoint (self):
        ppt = Indicator('ppt', False, name='PowerPoint Presentation',
            description='Contains a PowerPoint Document stream, very likely to be a Microsoft PowerPoint Presentation.')
        self.indicators.append(ppt)
        if self.ole.exists('PowerPoint Document'):
            ppt.value = True

    def check_visio (self):
        visio = Indicator('visio', False, name='Visio Drawing',
            description='Contains a VisioDocument stream, very likely to be a Microsoft Visio Drawing.')
        self.indicators.append(visio)
        if self.ole.exists('VisioDocument'):
            visio.value = True

    def check_ObjectPool (self):
        objpool = Indicator('ObjectPool', False, name='ObjectPool',
            description='Contains an ObjectPool stream, very likely to contain embedded OLE objects or files.')
        self.indicators.append(objpool)
        if self.ole.exists('ObjectPool'):
            objpool.value = True


    def check_flash (self):
        flash = Indicator('flash', 0, _type=int, name='Flash objects',
            description='Number of embedded Flash objects (SWF files) detected in OLE streams. Not 100% accurate, there may be false positives.')
        self.indicators.append(flash)
        for stream in self.ole.listdir():
            data = self.ole.openstream(stream).read()
            found = detect_flash(data)
            # just add to the count of Flash objects:
            flash.value += len(found)
            #print stream, found


#=== MAIN =================================================================

def main():
    usage = 'usage: %prog [options] <file>'
    parser = optparse.OptionParser(usage=__doc__ + '\n' + usage)
##    parser.add_option('-o', '--ole', action='store_true', dest='ole', help='Parse an OLE file (e.g. Word, Excel) to look for SWF in each stream')

    (options, args) = parser.parse_args()

    # Print help if no argurments are passed
    if len(args) == 0:
        parser.print_help()
        return

    for filename in args:
        print '\nFilename:', filename
        oleid = OleID(filename)
        indicators = oleid.check()

        #TODO: add description
        #TODO: highlight suspicious indicators
        t = prettytable.PrettyTable(['Indicator', 'Value'])
        t.align = 'l'
        t.max_width = 39
        #t.border = False

        for indicator in indicators:
            #print '%s: %s' % (indicator.name, indicator.value)
            t.add_row((indicator.name, indicator.value))

        print t

if __name__ == '__main__':
    main()
