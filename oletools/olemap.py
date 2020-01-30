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

# olemap is copyright (c) 2015-2019 Philippe Lagadec (http://www.decalage.info)
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
#                      - improved MiniFAT display with tablestream
# 2017-03-21       PL: - added header display
#                      - added options --header, --fat and --minifat
# 2017-03-22       PL: - added extra data detection, completed header display
# 2017-03-23       PL: - only display the header by default
#                      - added option --exdata to display extra data in hex
# 2018-08-28 v0.54 PL: - olefile is now a dependency
# 2019-07-10 v0.55 PL: - fixed display of OLE header CLSID (issue #394)

__version__ = '0.55'

#------------------------------------------------------------------------------
# TODO:

# === IMPORTS ================================================================

import sys, os, optparse, binascii

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

import olefile
from oletools.thirdparty.tablestream import tablestream
from oletools.thirdparty.xglob import xglob
from oletools.ezhexviewer import hexdump3

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


def show_header(ole, extra_data=False):
    print("OLE HEADER:")
    t = tablestream.TableStream([24, 16, 79-(4+24+16)], header_row=['Attribute', 'Value', 'Description'])
    t.write_row(['OLE Signature (hex)', binascii.b2a_hex(ole.header_signature).upper(), 'Should be D0CF11E0A1B11AE1'])
    t.write_row(['Header CLSID', ole.header_clsid, 'Should be empty (0)'])
    t.write_row(['Minor Version', '%04X' % ole.minor_version, 'Should be 003E'])
    t.write_row(['Major Version', '%04X' % ole.dll_version, 'Should be 3 or 4'])
    t.write_row(['Byte Order', '%04X' % ole.byte_order, 'Should be FFFE (little endian)'])
    t.write_row(['Sector Shift', '%04X' % ole.sector_shift, 'Should be 0009 or 000C'])
    t.write_row(['# of Dir Sectors', ole.num_dir_sectors, 'Should be 0 if major version is 3'])
    t.write_row(['# of FAT Sectors', ole.num_fat_sectors, ''])
    t.write_row(['First Dir Sector', '%08X' % ole.first_dir_sector, '(hex)'])
    t.write_row(['Transaction Sig Number', ole.transaction_signature_number, 'Should be 0'])
    t.write_row(['MiniStream cutoff', ole.mini_stream_cutoff_size, 'Should be 4096 bytes'])
    t.write_row(['First MiniFAT Sector', '%08X' % ole.first_mini_fat_sector, '(hex)'])
    t.write_row(['# of MiniFAT Sectors', ole.num_mini_fat_sectors, ''])
    t.write_row(['First DIFAT Sector', '%08X' % ole.first_difat_sector, '(hex)'])
    t.write_row(['# of DIFAT Sectors', ole.num_difat_sectors, ''])
    t.close()
    print('')
    print("CALCULATED ATTRIBUTES:")
    t = tablestream.TableStream([24, 16, 79-(4+24+16)], header_row=['Attribute', 'Value', 'Description'])
    t.write_row(['Sector Size (bytes)', ole.sector_size, 'Should be 512 or 4096 bytes'])
    t.write_row(['Actual File Size (bytes)', ole._filesize, 'Real file size on disk'])
    num_sectors_per_fat_sector = ole.sector_size/4
    num_sectors_in_fat = num_sectors_per_fat_sector * ole.num_fat_sectors
    # Need to add one sector for the header:
    max_filesize_fat = (num_sectors_in_fat + 1) * ole.sector_size
    t.write_row(['Max File Size in FAT', max_filesize_fat, 'Max file size covered by FAT'])
    if ole._filesize > max_filesize_fat:
        extra_size_beyond_fat = ole._filesize - max_filesize_fat
        color = 'red'
    else:
        extra_size_beyond_fat = 0
        color = None
    t.write_row(['Extra data beyond FAT', extra_size_beyond_fat, 'Only if file is larger than FAT coverage'],
                colors=[color, color, color])
    # Find the last used sector:
    # By default, it's the last sector in the FAT
    last_used_sector = len(ole.fat)-1
    for i in range(len(ole.fat)-1, 0, -1):
        last_used_sector = i
        if ole.fat[i] != olefile.FREESECT:
            break
    # Extra data would start at the next sector
    offset_extra_data = ole.sectorsize * (last_used_sector + 2)
    t.write_row(['Extra data offset in FAT', '%08X' % offset_extra_data, 'Offset of the 1st free sector at end of FAT'])
    extra_data_size = ole._filesize - offset_extra_data
    color = 'red' if extra_data_size > 0 else None
    t.write_row(['Extra data size', extra_data_size, 'Size of data starting at the 1st free sector at end of FAT'],
                colors=[color, color, color])
    t.close()
    print('')

    if extra_data:
        # hex dump of extra data
        print('HEX DUMP OF EXTRA DATA:\n')
        if extra_data_size <= 0:
            print('No extra data found at end of file.')
        else:
            ole.fp.seek(offset_extra_data)
            # read until end of file:
            exdata = ole.fp.read()
            assert len(exdata) == extra_data_size
            print(hexdump3(exdata, length=16, startindex=offset_extra_data))
        print('')


def show_fat(ole):
    print('FAT:')
    t = tablestream.TableStream([8, 12, 8, 8], header_row=['Sector #', 'Type', 'Offset', 'Next #'])
    for i in range(len(ole.fat)):
        fat_value = ole.fat[i]
        fat_type = FAT_TYPES.get(fat_value, '<Data>')
        color_type = FAT_COLORS.get(fat_value, FAT_COLORS['default'])
        # compute offset based on sector size:
        offset = ole.sectorsize * (i + 1)
        # print '%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value)
        t.write_row(['%8X' % i, fat_type, '%08X' % offset, '%8X' % fat_value],
                    colors=[None, color_type, None, None])
    t.close()
    print('')


def show_minifat(ole):
    print('MiniFAT:')
    # load MiniFAT if it wasn't already done:
    ole.loadminifat()
    t = tablestream.TableStream([8, 12, 8, 8], header_row=['Sector #', 'Type', 'Offset', 'Next #'])
    for i in range(len(ole.minifat)):
        fat_value = ole.minifat[i]
        fat_type = FAT_TYPES.get(fat_value, 'Data')
        color_type = FAT_COLORS.get(fat_value, FAT_COLORS['default'])
        # TODO: compute offset
        # print('%8X: %-12s offset=%08X next=%8X' % (i, fat_type, 0, fat_value))
        t.write_row(['%8X' % i, fat_type, 'N/A', '%8X' % fat_value],
                    colors=[None, color_type, None, None])
    t.close()
    print('')

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
    parser.add_option("--header", action="store_true", dest="header",
                      help='Display the OLE header (default: yes)')
    parser.add_option("--fat", action="store_true", dest="fat",
                      help='Display the FAT (default: no)')
    parser.add_option("--minifat", action="store_true", dest="minifat",
                      help='Display the MiniFAT (default: no)')
    parser.add_option('-x', "--exdata", action="store_true", dest="extra_data",
                      help='Display a hex dump of extra data at end of file')

    # TODO: add logfile option

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print(BANNER)
        print(__doc__)
        parser.print_help()
        sys.exit()

    # if no display option is provided, set defaults:
    default_options = False
    if not (options.header or options.fat or options.minifat):
        options.header = True
        # options.fat = True
        # options.minifat = True
        default_options = True

    # print banner with version
    print(BANNER)

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

        if options.header:
            show_header(ole, extra_data=options.extra_data)
        if options.fat:
            show_fat(ole)
        if options.minifat:
            show_minifat(ole)

        ole.close()

    # if no display option is provided, print a tip:
    if default_options:
        print('To display the FAT or MiniFAT structures, use options --fat or --minifat, and -h for help.')


if __name__ == '__main__':
    main()
