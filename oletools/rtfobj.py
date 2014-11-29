#!/usr/bin/env python
"""
rtfobj.py - Philippe Lagadec 2013-04-02

rtfobj is a Python module to extract embedded objects from RTF files, such as
OLE ojects. It can be used as a Python library or a command-line tool.

Usage: rtfobj.py <file.rtf>

rtfobj project website: http://www.decalage.info/python/rtfobj

rtfobj is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# rtfobj is copyright (c) 2012-2014, Philippe Lagadec (http://www.decalage.info)
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

__version__ = '0.02'

#------------------------------------------------------------------------------
# TODO:
# - improve regex pattern for better performance?
# - allow semicolon within hex, as found in  this sample:
#   http://contagiodump.blogspot.nl/2011/10/sep-28-cve-2010-3333-manuscript-with.html

#=== IMPORTS =================================================================

import re, sys, string, binascii


#=== CONSTANTS=================================================================

# REGEX pattern to extract embedded OLE objects in hexadecimal format:
# alphanum digit: [0-9A-Fa-f]
# hex char = two alphanum digits: [0-9A-Fa-f]{2}
# several hex chars, at least 4: (?:[0-9A-Fa-f]{2}){4,}
# at least 4 hex chars, followed by whitespace or CR/LF: (?:[0-9A-Fa-f]{2}){4,}\s*
PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'
# improved pattern, allowing semicolons within hex:
#PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'

# a dummy translation table for str.translate, which does not change anythying:
TRANSTABLE_NOCHANGE = string.maketrans('', '')


#=== FUNCTIONS =================================================================

def rtf_iter_objects (filename, min_size=32):
    """
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    data = open(filename, 'rb').read()
    for m in re.finditer(PATTERN, data):
        found = m.group(0)
        # remove all whitespace and line feeds:
        #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
        found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
        found = binascii.unhexlify(found)
        #print repr(found)
        if len(found)>min_size:
            yield m.start(), found


#=== MAIN =================================================================

if __name__ == '__main__':
    if len(sys.argv)<2:
        sys.exit(__doc__)
    for index, data in rtf_iter_objects(sys.argv[1]):
        print 'found object size %d at index %08X' % (len(data), index)
        fname = 'object_%08X.bin' % index
        print 'saving to file %s' % fname
        open(fname, 'wb').write(data)
