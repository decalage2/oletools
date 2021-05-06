#!/usr/bin/env python
"""
oledir.py

oledir parses OLE files to display technical information about their directory
entries, including deleted/orphan streams/storages and unused entries.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

oledir is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE ==================================================================

# oledir is copyright (c) 2015-2019 Philippe Lagadec (http://www.decalage.info)
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

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2015-04-17 v0.01 PL: - first version
# 2015-04-21 v0.02 PL: - improved display with prettytable
# 2016-01-13 v0.03 PL: - replaced prettytable by tablestream, added colors
# 2016-07-20 v0.50 SL: - added Python 3 support
# 2016-08-09       PL: - fixed issue #77 (imports from thirdparty dir)
# 2017-03-08 v0.51 PL: - fixed absolute imports, added optparse
#                      - added support for zip files and wildcards
# 2018-04-11 v0.53 PL: - added table displaying storage tree and CLSIDs
# 2018-04-13       PL: - moved KNOWN_CLSIDS to common.clsid
# 2018-08-28 v0.54 PL: - olefile is now a dependency
# 2018-10-06           - colorclass is now a dependency

__version__ = '0.54'

#------------------------------------------------------------------------------
# TODO:
# TODO: show FAT/MiniFAT
# TODO: show errors when reading streams

# === IMPORTS ================================================================

import sys, os, optparse

import olefile
import colorclass

# On Windows, colorclass needs to be enabled:
if os.name == 'nt':
    colorclass.Windows.enable(auto_colors=True)

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _parent_dir)
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty.tablestream import tablestream
from oletools.thirdparty.xglob import xglob
from oletools.common.clsid import KNOWN_CLSIDS

# === CONSTANTS ==============================================================

BANNER = 'oledir %s - http://decalage.info/python/oletools' % __version__

STORAGE_NAMES = {
    olefile.STGTY_EMPTY:     'Empty',
    olefile.STGTY_STORAGE:   'Storage',
    olefile.STGTY_STREAM:    'Stream',
    olefile.STGTY_LOCKBYTES: 'ILockBytes',
    olefile.STGTY_PROPERTY:  'IPropertyStorage',
    olefile.STGTY_ROOT:      'Root',
}

STORAGE_COLORS = {
    olefile.STGTY_EMPTY:     'green',
    olefile.STGTY_STORAGE:   'cyan',
    olefile.STGTY_STREAM:    'yellow',
    olefile.STGTY_LOCKBYTES: 'magenta',
    olefile.STGTY_PROPERTY:  'magenta',
    olefile.STGTY_ROOT:      'cyan',
}

STATUS_COLORS = {
    'unused':   'green',
    '<Used>':   'yellow',
    'ORPHAN':   'red',
}


# === FUNCTIONS ==============================================================

def sid_display(sid):
    if sid == olefile.NOSTREAM:
        return '-'  # None
    else:
        return sid

def clsid_display(clsid):
    clsid_upper = clsid.upper()
    if clsid_upper in KNOWN_CLSIDS:
        clsid += '\n%s' % KNOWN_CLSIDS[clsid_upper]
    color = 'yellow'
    if 'CVE' in clsid:
        color = 'red'
    return (clsid, color)

# === MAIN ===================================================================

def main():
    usage = 'usage: oledir [options] <filename> [filename2 ...]'
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

    if os.name == 'nt':
        colorclass.Windows.enable(auto_colors=True, reset_atexit=True)

    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
                                                      zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        full_name = '%s in %s' % (filename, container) if container else filename
        print('OLE directory entries in file %s:' % full_name)
        if data is not None:
            # data extracted from zip file
            ole = olefile.OleFileIO(data)
        else:
            # normal filename
            ole = olefile.OleFileIO(filename)
        # ole.dumpdirectory()

        # t = prettytable.PrettyTable(('id', 'Status', 'Type', 'Name', 'Left', 'Right', 'Child', '1st Sect', 'Size'))
        # t.align = 'l'
        # t.max_width['id'] = 4
        # t.max_width['Status'] = 6
        # t.max_width['Type'] = 10
        # t.max_width['Name'] = 10
        # t.max_width['Left'] = 5
        # t.max_width['Right'] = 5
        # t.max_width['Child'] = 5
        # t.max_width['1st Sect'] = 8
        # t.max_width['Size'] = 6

        table = tablestream.TableStream(column_width=[4, 6, 7, 22, 5, 5, 5, 8, 6],
            header_row=('id', 'Status', 'Type', 'Name', 'Left', 'Right', 'Child', '1st Sect', 'Size'),
            style=tablestream.TableStyleSlim)

        # TODO: read ALL the actual directory entries from the directory stream, because olefile does not!
        # TODO: OR fix olefile!
        # TODO: olefile should store or give access to the raw direntry data on demand
        # TODO: oledir option to hexdump the raw direntries
        # TODO: olefile should be less picky about incorrect directory structures

        for id in range(len(ole.direntries)):
            d = ole.direntries[id]
            if d is None:
                # this direntry is not part of the tree: either unused or an orphan
                d = ole._load_direntry(id) #ole.direntries[id]
                # print('%03d: %s *** ORPHAN ***' % (id, d.name))
                if d.entry_type == olefile.STGTY_EMPTY:
                    status = 'unused'
                else:
                    status = 'ORPHAN'
            else:
                # print('%03d: %s' % (id, d.name))
                status = '<Used>'
            if d.name.startswith('\x00'):
                # this may happen with unused entries, the name may be filled with zeroes
                name = ''
            else:
                # handle non-printable chars using repr(), remove quotes:
                name = repr(d.name)[1:-1]
            left  = sid_display(d.sid_left)
            right = sid_display(d.sid_right)
            child = sid_display(d.sid_child)
            entry_type = STORAGE_NAMES.get(d.entry_type, 'Unknown')
            etype_color = STORAGE_COLORS.get(d.entry_type, 'red')
            status_color = STATUS_COLORS.get(status, 'red')

            # print('      type=%7s sid_left=%s sid_right=%s sid_child=%s'
            #       %(entry_type, left, right, child))
            # t.add_row((id, status, entry_type, name, left, right, child, hex(d.isectStart), d.size))
            table.write_row((id, status, entry_type, name, left, right, child, '%X' % d.isectStart, d.size),
                colors=(None, status_color, etype_color, None, None, None, None, None, None))

        table = tablestream.TableStream(column_width=[4, 28, 6, 38],
            header_row=('id', 'Name', 'Size', 'CLSID'),
            style=tablestream.TableStyleSlim)
        rootname = ole.get_rootentry_name()
        entry_id = 0
        clsid = ole.root.clsid
        clsid_text, clsid_color = clsid_display(clsid)
        table.write_row((entry_id, rootname, '-', clsid_text),
                        colors=(None, 'cyan', None, clsid_color))
        for entry in sorted(ole.listdir(storages=True)):
            name = entry[-1]
            # handle non-printable chars using repr(), remove quotes:
            name = repr(name)[1:-1]
            name_color = None
            if ole.get_type(entry) in (olefile.STGTY_STORAGE, olefile.STGTY_ROOT):
                name_color = 'cyan'
            indented_name = '  '*(len(entry)-1) + name
            entry_id = ole._find(entry)
            try:
                size = ole.get_size(entry)
            except:
                size = '-'
            clsid = ole.getclsid(entry)
            clsid_text, clsid_color = clsid_display(clsid)
            table.write_row((entry_id, indented_name, size, clsid_text),
                            colors=(None, name_color, None, clsid_color))


        ole.close()
        # print t


if __name__ == '__main__':
    main()
