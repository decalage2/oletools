#!/usr/bin/env python
"""
oletimes.py - Philippe Lagadec 2013-07-24

oletimes is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract creation and modification times of all streams and storages
in the OLE file.

Usage: oletimes.py <file>

oletimes project website: http://www.decalage.info/python/oletimes

oletimes is part of the python-oletools package:
http://www.decalage.info/python/oletools

oletimes is copyright (c) 2013, Philippe Lagadec (http://www.decalage.info)
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
# 2013-07-24 v0.01 PL

#------------------------------------------------------------------------------
# TODO:
# + optparse
# + nicer output
# - CSV output
# - option to only show available timestamps (by default?)

import sys
from thirdparty.OleFileIO_PL import OleFileIO_PL


ole = OleFileIO_PL.OleFileIO(sys.argv[1])

print'- Root mtime=%s ctime=%s' % (ole.root.getmtime(), ole.root.getctime())

for obj in ole.listdir(streams=True, storages=True):
    print '- %s: mtime=%s ctime=%s' % (repr('/'.join(obj)), ole.getmtime(obj), ole.getctime(obj))

ole.close()
