#!/usr/bin/env python

"""
record_base.py

Common stuff for ole files whose streams are a sequence of record structures.
This is the case for xls and ppt, so classes are bases for xls_parser.py and
ppt_parser.py .
"""

# === LICENSE =================================================================
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#  * Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
# AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
# IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
# ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
# LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
# CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
# SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
# INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
# ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
# POSSIBILITY OF SUCH DAMAGE.

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2017-11-30 v0.01 CH: - first version based on xls_parser

#------------------------------------------------------------------------------
# TODO:
# - read DocumentSummaryInformation first to get more info about streams
#   (maybe content type or so; identify streams that are never record-based)
# - think about integrating this with olefile itself

# -----------------------------------------------------------------------------
#  REFERENCES:
#  - [MS-XLS]: Excel Binary File Format (.xls) Structure Specification
#    https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.14).aspx
#  - Understanding the Excel .xls Binary File Format
#    https://msdn.microsoft.com/en-us/library/office/gg615597(v=office.14).aspx
#  - [MS-PPT]


import sys
import os.path
from io import SEEK_CUR
import logging

# little hack to allow absolute imports even if oletools is not installed.
# Copied from olevba.py
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))  # pylint: disable=invalid-name
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))             # pylint: disable=invalid-name
del _thismodule_dir
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)
del _parent_dir

from oletools.thirdparty import olefile


###############################################################################
# Helpers
###############################################################################


ENTRY_TYPE2STR = {
    olefile.STGTY_EMPTY: 'empty',
    olefile.STGTY_STORAGE: 'storage',
    olefile.STGTY_STREAM: 'stream',
    olefile.STGTY_LOCKBYTES: 'lock-bytes',
    olefile.STGTY_PROPERTY: 'property',
    olefile.STGTY_ROOT: 'root'
}


###############################################################################
# Base Classes
###############################################################################


class OleRecordFile(olefile.OleFileIO):
    """ an OLE compound file whose streams have (mostly) record structure

    'record structure' meaning that streams are a sequence of records. Records
    are structure with information about type and size in their first bytes
    and type-dependent data of given size after that.

    Subclass of OleFileIO!
    """

    @classmethod
    def stream_class_for_name(cls, stream_name):
        """ helper for iter_streams, must be overwritten in subclasses """
        return OleRecordStream    # this is an abstract class!

    def iter_streams(self):
        """ find all streams, including orphans """
        logging.debug('Finding streams in ole file')

        for sid, direntry in enumerate(self.direntries):
            is_orphan = direntry is None
            if is_orphan:
                # this direntry is not part of the tree --> unused or orphan
                direntry = self._load_direntry(sid)
            is_stream = direntry.entry_type == olefile.STGTY_STREAM
            logging.debug('direntry {:2d} {}: {}'.format(
                sid, '[orphan]' if is_orphan else direntry.name,
                'is stream of size {}'.format(direntry.size) if is_stream else
                'no stream ({})'.format(ENTRY_TYPE2STR[direntry.entry_type])))
            if is_stream:
                clz = self.stream_class_for_name(direntry.name)
                yield clz(self._open(direntry.isectStart, direntry.size),
                          None if is_orphan else direntry.name)


class OleRecordStream(object):
    """ a stream found in an OleRecordFile

    Always has a name and a size (both read-only). Has an OleFileStream handle.

    abstract base class
    """

    def __init__(self, stream, name):
        self.stream = stream
        self.name = name
        self.size = stream.size

    def read_record_head(self):
        """ read first few bytes of record to determine size and type

        Abstract base method, to be implemented in subclasses.

        returns (rec_type, rec_size, other) where other will be forwarded to
        record constructors
        """
        raise NotImplementedError('Abstract method '
                                  'OleRecordStream.read_record_head called')

    @classmethod
    def record_class_for_type(cls, rec_type):
        """ determine a class for given record type

        Only a base implementation. Create subclasses of OleRecordBase and
        return those when appropriate.

        returns (clz, force_read)
        """
        return OleRecordBase, False

    def iter_records(self, fill_data=False):
        """ yield all records in this stream

        Stream must be positioned at start of records (e.g. start of stream).
        """
        while True:
            # unpacking as in olevba._extract_vba
            pos = self.stream.tell()
            if pos >= self.size:
                break

            # read first few bytes, determine record type and size
            rec_type, rec_size, other = self.read_record_head()
            logging.debug('Record type {0} of size {1}'
                          .format(rec_type, rec_size))

            # determine what class to wrap this into
            rec_clz, force_read = self.record_class_for_type(rec_type)

            if fill_data or force_read:
                data = self.stream.read(rec_size)
                if len(data) != rec_size:
                    raise IOError('Not enough data in stream ({0} < {1})'
                                  .format(len(data), rec_size))
            else:
                self.stream.seek(rec_size, SEEK_CUR)
                data = None
            yield rec_clz(rec_type, rec_size, other, pos, data)

    def __str__(self):
        return '[{2} {0} (size {1})' \
               .format(self.name or '[orphan]', self.size,
                       self.__class__.__name__)


class OleRecordBase(object):
    """ a record found in an OleRecordStream

    always has a type and a size, also pos and data
    """

    # for subclasses with a fixed type
    TYPE = None

    # (max) size of subclasses
    MAX_SIZE = None
    SIZE = None

    def __init__(self, type, size, more_data, pos, data):
        """ create a record; more_data is discarded """
        if self.TYPE is not None and type != self.TYPE:
            raise ValueError('Wrong subclass {0} for type {1}'
                             .format(self.__class__.__name__, type))
        self.type = type
        if self.SIZE is not None and size != self.SIZE:
            raise ValueError('Wrong size {0} for record type {1}'
                             .format(size, type))
        elif self.MAX_SIZE is not None and size > self.MAX_SIZE:
            raise ValueError('Wrong size: {0} > MAX_SIZE for record type {1}'
                             .format(size, type))
        self.size = size
        self.pos = pos
        self.data = data
        self.parse(more_data)

    def parse(self, more_data):
        """ finish constructing this record

        Can save more_data from OleRecordStream.read_record_head and/or parse
        data (if it was read).

        Base implementation, does nothing. To be overwritten in subclasses.
        """
        pass

    def _type_str(self):
        """ helper for __str__, base implementation """
        return '{0} type {1}'.format(self.__class__.__name__, self.type)

    def __str__(self):
        """ create a short but informative textual representation of self """
        return '[' + self._type_str() + \
               ' (size {0} from {1})]'.format(self.size, self.pos)


###############################################################################
# TESTING
###############################################################################


def test(filenames, ole_file_class=OleRecordFile,
         must_parse=None):
    """ parse all given file names and print rough structure

    if an error occurs while parsing a stream of type in must_parse, the error
    will be raised. Otherwise a message is printed
    """
    logging.basicConfig(level=logging.DEBUG)
    if not filenames:
        logging.info('need file name[s]')
        return 2
    for filename in filenames:
        logging.info('checking file {0}'.format(filename))
        if not olefile.isOleFile(filename):
            logging.info('not an ole file - skip')
            continue
        ole = ole_file_class(filename)

        for stream in ole.iter_streams():
            logging.info(stream)
            try:
                for record in stream.iter_records():
                    logging.info('  {0}'.format(record))
            except Exception:
                if not must_parse:
                    raise
                elif isinstance(stream, must_parse):
                    raise
                else:
                    logging.info('  failed to parse', exc_info=True)
    return 0


if __name__ == '__main__':
    sys.exit(test(sys.argv[1:]))
