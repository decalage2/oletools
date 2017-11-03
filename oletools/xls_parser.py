""" Parse xls up to some point

Read storages, (sub-)streams, records from xls file
"""
#
# === LICENSE ==================================================================

# xls_parser is copyright (c) 2014-2017 Philippe Lagadec (http://www.decalage.info)
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
# 2017-11-02 v0.01 CH: - first version

__version__ = '0.1'

#------------------------------------------------------------------------------
# TODO:
# everything
#
#------------------------------------------------------------------------------
# REFERENCES:
# - [MS-XLS]: Excel Binary File Format (.xls) Structure Specification
#   https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.14).aspx
# - Understanding the Excel .xls Binary File Format
#   https://msdn.microsoft.com/en-us/library/office/gg615597(v=office.14).aspx
#
#--- IMPORTS ------------------------------------------------------------------

import sys

# little hack to allow absolute imports even if oletools is not installed.
# Copied from olevba.py
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty import olefile


class XlsFile(olefile.OleFileIO):
    """ specialization of an OLE compound file """

    def get_streams_recursive(self, storage=None):
        """ find all streams in all storages, depth-first """
        if storage is None:
            storage = self
        print('Finding streams in ole file')
        for st in storage.listdir(streams=True, storages=True):
            st_type = self.get_type(st)
            if st_type == olefile.STGTY_STREAM:      # a stream --> yield
                print('Checking stream {0}'.format(st))
                yield st
            elif st_type == olefile.STGTY_STORAGE:   # a storage --> recurse
                print('Recurse into storage {0}'.format(st))
                for entry in self.get_streams_recursive(st):
                    yield entry
            else:
                raise ValueError('unexpected type {0} for entry {1}'
                                 .format(st_type, st))


class XlsStream:
    """ specialization of an OLE (sub-)stream """
    pass


def test(filename):
    """ parse given file and print rough structure """
    pass

if __name__ == '__main__':
    """ parse all given file names and print rough structure """
    for filename in sys.argv[1:]:
        test(filename)
