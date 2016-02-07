#!/usr/bin/env python
"""
oledir.py

oledir parses OLE files to display technical information about its directory
entries, including deleted/orphan streams/storages and unused entries.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

oledir is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE ==================================================================

# oledir is copyright (c) 2015-2016 Philippe Lagadec (http://www.decalage.info)
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
# 2015-04-17 v0.01 PL: - first version
# 2015-04-21 v0.02 PL: - improved display with prettytable
# 2016-01-13 v0.03 PL: - replaced prettytable by tablestream, added colors

__version__ = '0.03'

#------------------------------------------------------------------------------
# TODO:
# TODO: show FAT/MiniFAT
# TODO: show errors when reading streams

# === IMPORTS ================================================================

import sys, os
from thirdparty.olefile import olefile
# from thirdparty.prettytable import prettytable
from thirdparty.tablestream import tablestream
from thirdparty.colorclass import colorclass


def sid_display(sid):
    if sid == olefile.NOSTREAM:
        return '-' #None
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

STORAGE_COLORS = {
    olefile.STGTY_EMPTY:     'green',
    olefile.STGTY_STORAGE:   'blue',
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

# === MAIN ===================================================================

if __name__ == '__main__':
    # print banner with version
    print 'oledir %s - http://decalage.info/python/oletools' % __version__

    if os.name == 'nt':
        colorclass.Windows.enable(auto_colors=True, reset_atexit=True)

    fname = sys.argv[1]
    print('OLE directory entries in file %s:' % fname)
    ole = olefile.OleFileIO(fname)
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

    for id in xrange(len(ole.direntries)):
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
    ole.close()
    # print t


