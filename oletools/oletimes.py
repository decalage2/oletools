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

# oletimes is copyright (c) 2013-2019, Philippe Lagadec (http://www.decalage.info)
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
# 2016-07-20 v0.50 SL: - added Python 3 support
# 2016-09-05       PL: - added main entry point for setup.py
# 2017-05-03 v0.51 PL: - fixed absolute imports (issue #141)
# 2017-05-04       PL: - added optparse and xglob (issue #141)
# 2018-09-11 v0.54 PL: - olefile is now a dependency

__version__ = '0.54'

#------------------------------------------------------------------------------
# TODO:
# + nicer output: table with fixed columns, datetime, etc
# + CSV output
# + option to only show available timestamps (by default?)

#=== IMPORTS =================================================================

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

import olefile
from oletools.thirdparty import xglob
from oletools.thirdparty.prettytable import prettytable


# === FUNCTIONS ==============================================================

def dt2str(dt):
    """
    Convert a datetime object to a string for display, without microseconds

    :param dt: datetime.datetime object, or None
    :return: str, or None
    """
    if dt is None:
        return None
    dt = dt.replace(microsecond=0)
    return str(dt)


def process_ole(ole):
    t = prettytable.PrettyTable(['Stream/Storage name', 'Modification Time', 'Creation Time'])
    t.align = 'l'
    t.max_width = 26
    t.add_row(('Root', dt2str(ole.root.getmtime()), dt2str(ole.root.getctime())))
    for obj in ole.listdir(streams=True, storages=True):
        t.add_row((repr('/'.join(obj)), dt2str(ole.getmtime(obj)), dt2str(ole.getctime(obj))))
    print(t)


# === MAIN ===================================================================

def main():
    # print banner with version
    print('oletimes %s - http://decalage.info/python/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')

    usage = 'usage: oletimes [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    parser.add_option("-r", action="store_true", dest="recursive",
                      help='find files recursively in subdirectories.')
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
                      help='if the file is a zip archive, open all files from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
                      help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')

    # TODO: add logfile option
    # parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
    #                         help="logging level debug/info/warning/error/critical (default=%default)")

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print(__doc__)
        parser.print_help()
        sys.exit()

    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
                                                      zip_password=options.zip_password, zip_fname=options.zip_fname):
        # TODO: handle xglob errors
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        full_name = '%s in %s' % (filename, container) if container else filename
        print("=" * 79)
        print('FILE: %s\n' % full_name)
        if data is not None:
            # data extracted from zip file
            ole = olefile.OleFileIO(data)
        else:
            # normal filename
            ole = olefile.OleFileIO(filename)
        process_ole(ole)
        ole.close()

if __name__ == '__main__':
    main()
