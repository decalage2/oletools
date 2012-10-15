#!/usr/bin/env python
"""
pyxswf.py - Philippe Lagadec 2012-09-17

pyxswf is a script to detect, extract and analyze Flash objects (SWF) that may
be embedded in files such as MS Office documents (e.g. Word, Excel),
which is especially useful for malware analysis.
pyxswf is an improved version of xxxswf.py published by Alexander Hanel on
http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html
Compared to xxxswf, it can extract streams from MS Office documents by parsing
their OLE structure properly, which is necessary when streams are fragmented.
Stream fragmentation is a known obfuscation technique, as explained on
http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/

pyxswf project website: http://www.decalage.info/python/pyxswf

pyxswf is copyright (c) 2012, Philippe Lagadec (http://www.decalage.info)
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

 * Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
 * Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

__version__ = '0.01'

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-09-17 v0.01 PL: - first version

#------------------------------------------------------------------------------
# TODO:
# - check if file is OLE
# - support -r

import optparse, sys, os
from thirdparty.xxxswf import xxxswf
from thirdparty.OleFileIO_PL import OleFileIO_PL

def main():
    # Scenarios:
    # Scan file for SWF(s)
    # Scan file for SWF(s) and extract them
    # Scan file for SWF(s) and scan them with Yara
    # Scan file for SWF(s), extract them and scan with Yara
    # Scan directory recursively for files that contain SWF(s)
    # Scan directory recursively for files that contain SWF(s) and extract them

    usage = 'usage: %prog [options] <file.bad>'
    parser = optparse.OptionParser(usage=usage)
    parser.add_option('-x', '--extract', action='store_true', dest='extract', help='Extracts the embedded SWF(s), names it MD5HASH.swf & saves it in the working dir. No addition args needed')
    parser.add_option('-y', '--yara', action='store_true', dest='yara', help='Scans the SWF(s) with yara. If the SWF(s) is compressed it will be deflated. No addition args needed')
    parser.add_option('-s', '--md5scan', action='store_true', dest='md5scan', help='Scans the SWF(s) for MD5 signatures. Please see func checkMD5 to define hashes. No addition args needed')
    parser.add_option('-H', '--header', action='store_true', dest='header', help='Displays the SWFs file header. No addition args needed')
    parser.add_option('-d', '--decompress', action='store_true', dest='decompress', help='Deflates compressed SWFS(s)')
    parser.add_option('-r', '--recdir', dest='PATH', type='string', help='Will recursively scan a directory for files that contain SWFs. Must provide path in quotes')
    parser.add_option('-c', '--compress', action='store_true', dest='compress', help='Compresses the SWF using Zlib')

    parser.add_option('-o', '--ole', action='store_true', dest='ole', help='Parse an OLE file (e.g. Word, Excel) to look for SWF in each stream')


    (options, args) = parser.parse_args()

    # Print help if no argurments are passed
    if len(args) == 0:
        parser.print_help()
        return

    if options.ole:
        for filename in args:
            ole = OleFileIO_PL.OleFileIO(filename)
            for direntry in ole.direntries:
                if direntry is not None and direntry.entry_type == OleFileIO_PL.STGTY_STREAM:
                    f = ole._open(direntry.isectStart, direntry.size)
                    # check if data contains the SWF magic: FWS or CWS
                    data = f.getvalue()
                    if 'FWS' in data or 'CWS' in data:
                        print 'OLE stream: %s' % repr(direntry.name)
                        # call xxxswf to scan or extract Flash files:
                        xxxswf.disneyland(f, direntry.name, options)
                    f.close()
            ole.close()
    else:
        xxxswf.main()

if __name__ == '__main__':
    main()
