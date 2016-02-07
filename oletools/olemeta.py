#!/usr/bin/env python
"""
olemeta.py

olemeta is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract all standard properties present in the OLE file.

Usage: olemeta.py <file>

olemeta project website: http://www.decalage.info/python/olemeta

olemeta is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# olemeta is copyright (c) 2013-2015, Philippe Lagadec (http://www.decalage.info)
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
# 2013-07-24 v0.01 PL: - first version
# 2014-11-29 v0.02 PL: - use olefile instead of OleFileIO_PL
#                      - improved usage display
# 2015-12-29 v0.03 PL: - only display properties present in the file

__version__ = '0.03'

#------------------------------------------------------------------------------
# TODO:
# + optparse
# + nicer output: table with fixed columns, datetime, etc
# + CSV output
# + option to only show available properties (by default)

#=== IMPORTS =================================================================

import sys, codecs
import thirdparty.olefile as olefile
from thirdparty.tablestream import tablestream


#=== MAIN =================================================================

try:
    ole = olefile.OleFileIO(sys.argv[1])
except IndexError:
    sys.exit(__doc__)

# parse and display metadata:
meta = ole.get_metadata()

# console output with UTF8 encoding:
console_utf8 = codecs.getwriter('utf8')(sys.stdout)

# TODO: move similar code to a function

print('Properties from the SummaryInformation stream:')
t = tablestream.TableStream([21, 30], header_row=['Property', 'Value'], outfile=console_utf8)
for prop in meta.SUMMARY_ATTRIBS:
    value = getattr(meta, prop)
    if value is not None:
        # TODO: pretty printing for strings, dates, numbers
        # TODO: better unicode handling
        # print('- %s: %s' % (prop, value))
        if isinstance(value, unicode):
            # encode to UTF8, avoiding errors
            value = value.encode('utf-8', errors='replace')
        else:
            value = str(value)
        t.write_row([prop, value], colors=[None, 'yellow'])
t.close()
print ''

print('Properties from the DocumentSummaryInformation stream:')
t = tablestream.TableStream([21, 30], header_row=['Property', 'Value'], outfile=console_utf8)
for prop in meta.DOCSUM_ATTRIBS:
    value = getattr(meta, prop)
    if value is not None:
        # TODO: pretty printing for strings, dates, numbers
        # TODO: better unicode handling
        # print('- %s: %s' % (prop, value))
        if isinstance(value, unicode):
            # encode to UTF8, avoiding errors
            value = value.encode('utf-8', errors='replace')
        else:
            value = str(value)
        t.write_row([prop, value], colors=[None, 'yellow'])
t.close()

ole.close()
