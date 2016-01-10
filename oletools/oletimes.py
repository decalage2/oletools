#!/usr/bin/env python
"""
oletimes.py

oletimes is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract creation and modification times of all streams and storages
in the OLE file.

Usage: oletimes.py <file>

oletimes project website: http://www.decalage.info/python/oletimes

oletimes is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# oletimes is copyright (c) 2013-2015, Philippe Lagadec (http://www.decalage.info)
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
# 2014-11-30 v0.03 PL: - improved output with prettytable

__version__ = '0.03'

#------------------------------------------------------------------------------
# TODO:
# + optparse
# + nicer output: table with fixed columns, datetime, etc
# + CSV output
# + option to only show available timestamps (by default?)

#=== IMPORTS =================================================================

import sys, datetime
import thirdparty.olefile as olefile
from thirdparty.prettytable import prettytable


#=== MAIN =================================================================

try:
    ole = olefile.OleFileIO(sys.argv[1])
except IndexError:
    sys.exit(__doc__)

def dt2str (dt):
    """
    Convert a datetime object to a string for display, without microseconds

    :param dt: datetime.datetime object, or None
    :return: str, or None
    """
    if dt is None:
        return None
    dt = dt.replace(microsecond = 0)
    return str(dt)

t = prettytable.PrettyTable(['Stream/Storage name', 'Modification Time', 'Creation Time'])
t.align = 'l'
t.max_width = 26
#t.border = False

#print'- Root mtime=%s ctime=%s' % (ole.root.getmtime(), ole.root.getctime())
t.add_row(('Root', dt2str(ole.root.getmtime()), dt2str(ole.root.getctime())))

for obj in ole.listdir(streams=True, storages=True):
    #print '- %s: mtime=%s ctime=%s' % (repr('/'.join(obj)), ole.getmtime(obj), ole.getctime(obj))
    t.add_row((repr('/'.join(obj)), dt2str(ole.getmtime(obj)), dt2str(ole.getctime(obj))))

print t

ole.close()
