#!/usr/bin/env python
"""
pyxswf.py

pyxswf is a script to detect, extract and analyze Flash objects (SWF) that may
be embedded in files such as MS Office documents (e.g. Word, Excel),
which is especially useful for malware analysis.

pyxswf is an extension to xxxswf.py published by Alexander Hanel on
http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html
Compared to xxxswf, it can extract streams from MS Office documents by parsing
their OLE structure properly (-o option), which is necessary when streams are
fragmented.
Stream fragmentation is a known obfuscation technique, as explained on
http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/

It can also extract Flash objects from RTF documents, by parsing embedded
objects encoded in hexadecimal format (-f option).

pyxswf project website: http://www.decalage.info/python/pyxswf

pyxswf is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# pyxswf is copyright (c) 2012-2019, Philippe Lagadec (http://www.decalage.info)
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
# 2012-09-17 v0.01 PL: - first version
# 2012-11-09 v0.02 PL: - added RTF embedded objects extraction
# 2014-11-29 v0.03 PL: - use olefile instead of OleFileIO_PL
#                      - improved usage display with -h
# 2016-09-06 v0.50 PL: - updated to match the rtfobj API
# 2016-10-25       PL: - fixed print for Python 3
# 2016-11-01       PL: - replaced StringIO by BytesIO for Python 3
# 2018-09-11 v0.54 PL: - olefile is now a dependency

__version__ = '0.54'

#------------------------------------------------------------------------------
# TODO:
# + update xxxswf to latest version
# + add support for LZMA-compressed flash files (ZWS header)
#   references: http://blog.malwaretracker.com/2014/01/cve-2013-5331-evaded-av-by-using.html
#   http://code.metager.de/source/xref/adobe/flash/crossbridge/tools/swf-info.py
#   http://room32.dyndns.org/forums/showthread.php?766-SWFCompression
#   sample code: http://room32.dyndns.org/SWFCompression.py
# - check if file is OLE
# - support -r


#=== IMPORTS =================================================================

import optparse, sys, os
from . import rtfobj
from io import BytesIO
from .thirdparty.xxxswf import xxxswf
import olefile


#=== MAIN =================================================================

def main():
    # print banner with version
    print ('pyxswf %s - http://decalage.info/python/oletools' % __version__)
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')
    # Scenarios:
    # Scan file for SWF(s)
    # Scan file for SWF(s) and extract them
    # Scan file for SWF(s) and scan them with Yara
    # Scan file for SWF(s), extract them and scan with Yara
    # Scan directory recursively for files that contain SWF(s)
    # Scan directory recursively for files that contain SWF(s) and extract them

    usage = 'usage: %prog [options] <file.bad>'
    parser = optparse.OptionParser(usage=__doc__ + '\n' + usage)
    parser.add_option('-x', '--extract', action='store_true', dest='extract', help='Extracts the embedded SWF(s), names it MD5HASH.swf & saves it in the working dir. No addition args needed')
    parser.add_option('-y', '--yara', action='store_true', dest='yara', help='Scans the SWF(s) with yara. If the SWF(s) is compressed it will be deflated. No addition args needed')
    parser.add_option('-s', '--md5scan', action='store_true', dest='md5scan', help='Scans the SWF(s) for MD5 signatures. Please see func checkMD5 to define hashes. No addition args needed')
    parser.add_option('-H', '--header', action='store_true', dest='header', help='Displays the SWFs file header. No addition args needed')
    parser.add_option('-d', '--decompress', action='store_true', dest='decompress', help='Deflates compressed SWFS(s)')
    parser.add_option('-r', '--recdir', dest='PATH', type='string', help='Will recursively scan a directory for files that contain SWFs. Must provide path in quotes')
    parser.add_option('-c', '--compress', action='store_true', dest='compress', help='Compresses the SWF using Zlib')

    parser.add_option('-o', '--ole', action='store_true', dest='ole', help='Parse an OLE file (e.g. Word, Excel) to look for SWF in each stream')
    parser.add_option('-f', '--rtf', action='store_true', dest='rtf', help='Parse an RTF file to look for SWF in each embedded object')


    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        parser.print_help()
        return

    # OLE MODE:
    if options.ole:
        for filename in args:
            ole = olefile.OleFileIO(filename)
            for direntry in ole.direntries:
                if direntry is not None and direntry.entry_type == olefile.STGTY_STREAM:
                    f = ole._open(direntry.isectStart, direntry.size)
                    # check if data contains the SWF magic: FWS or CWS
                    data = f.getvalue()
                    if b'FWS' in data or b'CWS' in data:
                        print('OLE stream: %s' % repr(direntry.name))
                        # call xxxswf to scan or extract Flash files:
                        xxxswf.disneyland(f, direntry.name, options)
                    f.close()
            ole.close()

    # RTF MODE:
    elif options.rtf:
        for filename in args:
            for index, orig_len, data in rtfobj.rtf_iter_objects(filename):
                if b'FWS' in data or b'CWS' in data:
                    print('RTF embedded object size %d at index %08X' % (len(data), index))
                    f = BytesIO(data)
                    name = 'RTF_embedded_object_%08X' % index
                    # call xxxswf to scan or extract Flash files:
                    xxxswf.disneyland(f, name, options)

    else:
        xxxswf.main()

if __name__ == '__main__':
    main()
