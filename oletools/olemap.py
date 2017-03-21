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

# olemap is copyright (c) 2015-2017 Philippe Lagadec (http://www.decalage.info)
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
# 2016-07-20 v0.50 SL: - added Python 3 support
# 2016-09-05       PL: - added main entry point for setup.py
# 2017-03-20 v0.51 PL: - fixed absolute imports, added optparse
#                      - added support for zip files and wildcards


__version__ = '0.51dev3'

#------------------------------------------------------------------------------
# TODO:

# === IMPORTS ================================================================

import sys, os, optparse

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty.olefile import olefile
from oletools.thirdparty.tablestream import tablestream
from oletools.thirdparty.xglob import xglob

# === CONSTANTS ==============================================================

BANNER = 'olemap %s - http://decalage.info/python/oletools' % __version__

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


# === FUNCTIONS ==============================================================

def sid_display(sid):
    if sid == olefile.NOSTREAM:
        return None
    else:
        return sid


# === MAIN ===================================================================

def main():
    usage = 'usage: olemap [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    parser.add_option("-r", action="store_true", dest="recursive",
                      help='find files recursively in subdirectories.')
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
                      help='if the file is a zip archive, open all files from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
                      help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')
    # parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
    #                         help="logging level debug/info/warning/error/critical (default=%default)")

    # TODO: add logfile option

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print(BANNER)
        print(__doc__)
        parser.print_help()
        sys.exit()

    # print banner with version
    print(BANNER)
    # print banner with version
    print('olemap %s - http://decalage.info/python/oletools' % __version__)

    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
                                                      zip_password=options.zip_password, zip_fname=options.zip_fname):
        # TODO: handle xglob errors
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        full_name = '%s in %s' % (filename, container) if container else filename
        print("-" * 79)
        print('FILE: %s\n' % full_name)
        if data is not None:
            # data extracted from zip file
            ole = olefile.OleFileIO(data)
        else:
            # normal filename
            ole = olefile.OleFileIO(filename)
        print('FAT:')
        t = tablestream.TableStream([8, 12, 8, 8], header_row=['Sector #', 'Type', 'Offset', 'Next #'])
        for i in range(ole.nb_sect):
            fat_value = ole.fat[i]
            fat_type = FAT_TYPES.get(fat_value, '<Data>')
            color_type = FAT_COLORS.get(fat_value, FAT_COLORS['default'])
            # compute offset based on sector size:
            offset = ole.sectorsize * (i+1)
            # print '%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value)
            t.write_row(['%8X' % i, fat_type, '%08X' % offset, '%8X' % fat_value],
                colors=[None, color_type, None, None])
        print('')

        print('MiniFAT:')
        # load MiniFAT if it wasn't already done:
        ole.loadminifat()
        for i in range(len(ole.minifat)):
            fat_value = ole.minifat[i]
            fat_type = FAT_TYPES.get(fat_value, 'Data')
            print('%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value))

        ole.close()

if __name__ == '__main__':
    main()
