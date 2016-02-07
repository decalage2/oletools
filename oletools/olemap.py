#!/usr/bin/env python
"""
olemap

olemap parses OLE files to display technical information about its structure.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

olemap is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE ==================================================================

# olemap is copyright (c) 2015 Philippe Lagadec (http://www.decalage.info)
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
# 2015-11-01 v0.01 PL: - first version
# 2016-01-13 v0.02 PL: - improved display with tablestream, added colors

__version__ = '0.02'

#------------------------------------------------------------------------------
# TODO:

# === IMPORTS ================================================================

import sys
from thirdparty.olefile import olefile
from thirdparty.tablestream import tablestream



def sid_display(sid):
    if sid == olefile.NOSTREAM:
        return None
    else:
        return sid

STORAGE_NAMES = {
    olefile.STGTY_EMPTY:     'Empty',
    olefile.STGTY_STORAGE:   'Storage',
    olefile.STGTY_STREAM:    'Stream',
    olefile.STGTY_LOCKBYTES: 'ILockBytes',
    olefile.STGTY_PROPERTY:  'IPropertyStorage',
    olefile.STGTY_ROOT:      'Root',
}

FAT_TYPES = {
    olefile.FREESECT:   "Free",
    olefile.ENDOFCHAIN: "End of Chain",
    olefile.FATSECT:    "FAT Sector",
    olefile.DIFSECT:    "DIFAT Sector"
    }

FAT_COLORS = {
    olefile.FREESECT:   "green",
    olefile.ENDOFCHAIN: "yellow",
    olefile.FATSECT:    "cyan",
    olefile.DIFSECT:    "blue",
    'default':          None,
    }


# === MAIN ===================================================================

if __name__ == '__main__':
    # print banner with version
    print 'olemap %s - http://decalage.info/python/oletools' % __version__

    fname = sys.argv[1]
    ole = olefile.OleFileIO(fname)

    print 'FAT:'
    t = tablestream.TableStream([8, 12, 8, 8], header_row=['Sector #', 'Type', 'Offset', 'Next #'])
    for i in xrange(ole.nb_sect):
        fat_value = ole.fat[i]
        fat_type = FAT_TYPES.get(fat_value, '<Data>')
        color_type = FAT_COLORS.get(fat_value, FAT_COLORS['default'])
        # compute offset based on sector size:
        offset = ole.sectorsize * (i+1)
        # print '%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value)
        t.write_row(['%8X' % i, fat_type, '%08X' % offset, '%8X' % fat_value],
            colors=[None, color_type, None, None])
    print ''

    print 'MiniFAT:'
    # load MiniFAT if it wasn't already done:
    ole.loadminifat()
    for i in xrange(len(ole.minifat)):
        fat_value = ole.minifat[i]
        fat_type = FAT_TYPES.get(fat_value, 'Data')
        print '%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value)

    ole.close()


