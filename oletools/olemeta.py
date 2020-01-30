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

# olemeta is copyright (c) 2013-2019, Philippe Lagadec (http://www.decalage.info)
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
# 2016-09-06 v0.50 PL: - added main entry point for setup.py
# 2016-10-25       PL: - fixed print for Python 3
# 2016-10-28       PL: - removed the UTF8 codec for console display
# 2017-04-26 v0.51 PL: - fixed absolute imports (issue #141)
# 2017-05-04       PL: - added optparse and xglob (issue #141)
# 2018-09-11 v0.54 PL: - olefile is now a dependency

__version__ = '0.54'

#------------------------------------------------------------------------------
# TODO:
# + nicer output: table with fixed columns, datetime, etc
# + CSV output
# + option to only show available properties (by default)
# + display codepage names

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
from oletools.thirdparty.tablestream import tablestream
from oletools.common.io_encoding import ensure_stdout_handles_unicode


#=== MAIN =================================================================

def process_ole(ole):
    # parse and display metadata:
    meta = ole.get_metadata()

    # console output with UTF8 encoding:
    ensure_stdout_handles_unicode()

    # TODO: move similar code to a function

    print('Properties from the SummaryInformation stream:')
    t = tablestream.TableStream([21, 30], header_row=['Property', 'Value'])
    for prop in meta.SUMMARY_ATTRIBS:
        value = getattr(meta, prop)
        if value is not None:
            # TODO: pretty printing for strings, dates, numbers
            # TODO: better unicode handling
            # print('- %s: %s' % (prop, value))
            # if isinstance(value, unicode):
            #     # encode to UTF8, avoiding errors
            #     value = value.encode('utf-8', errors='replace')
            # else:
            #     value = str(value)
            t.write_row([prop, value], colors=[None, 'yellow'])
    t.close()
    print('')

    print('Properties from the DocumentSummaryInformation stream:')
    t = tablestream.TableStream([21, 30], header_row=['Property', 'Value'])
    for prop in meta.DOCSUM_ATTRIBS:
        value = getattr(meta, prop)
        if value is not None:
            # TODO: pretty printing for strings, dates, numbers
            # TODO: better unicode handling
            # print('- %s: %s' % (prop, value))
            # if isinstance(value, unicode):
            #     # encode to UTF8, avoiding errors
            #     value = value.encode('utf-8', errors='replace')
            # else:
            #     value = str(value)
            t.write_row([prop, value], colors=[None, 'yellow'])
    t.close()


# === MAIN ===================================================================

def main():
    # print banner with version
    print('olemeta %s - http://decalage.info/python/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')

    usage = 'usage: olemeta [options] <filename> [filename2 ...]'
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
