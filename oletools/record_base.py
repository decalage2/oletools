#!/usr/bin/env python

"""
record_base.py

Common stuff for ole files whose streams are a sequence of record structures.
This is the case for xls and ppt, so classes are bases for xls_parser.py and
ppt_record_parser.py .
"""

# === LICENSE ==================================================================

# record_base is copyright (c) 2014-2024 Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
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

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-11-30 v0.01 CH: - first version based on xls_parser
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2019-01-30       PL: - fixed import to avoid mixing installed oletools
#                        and dev version
# 2019-05-24       CH: - use log_helper

__version__ = '0.60.2'

# -----------------------------------------------------------------------------
# TODO:
# - read DocumentSummaryInformation first to get more info about streams
#   (maybe content type or so; identify streams that are never record-based)
#   Or use oleid to avoid same functionality in several files
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

import olefile

# little hack to allow absolute imports even if oletools is not installed.
PARENT_DIR = os.path.normpath(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))
if PARENT_DIR not in sys.path:
    sys.path.insert(0, PARENT_DIR)
del PARENT_DIR
from oletools.common.log_helper import log_helper


###############################################################################
# Helpers
###############################################################################

OleFileIO = olefile.OleFileIO
STGTY_EMPTY     = olefile.STGTY_EMPTY      # 0
STGTY_STORAGE   = olefile.STGTY_STORAGE    # 1
STGTY_STREAM    = olefile.STGTY_STREAM     # 2
STGTY_LOCKBYTES = olefile.STGTY_LOCKBYTES  # 3
STGTY_PROPERTY  = olefile.STGTY_PROPERTY   # 4
STGTY_ROOT      = olefile.STGTY_ROOT       # 5
STGTY_SUBSTREAM = 10

ENTRY_TYPE2STR = {
    olefile.STGTY_EMPTY: 'empty',
    olefile.STGTY_STORAGE: 'storage',
    olefile.STGTY_STREAM: 'stream',
    olefile.STGTY_LOCKBYTES: 'lock-bytes',
    olefile.STGTY_PROPERTY: 'property',
    olefile.STGTY_ROOT: 'root',
    STGTY_SUBSTREAM: 'substream'
}


logger = log_helper.get_or_create_silent_logger('record_base')


def enable_olefile_logging():
    """ enable logging in olefile e.g., to get debug info from OleFileIO """
    olefile.enable_logging()


def enable_logging():
    """
    Enable logging for this module (disabled by default).

    For use by third-party libraries that import `record_base` as module.

    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    logger.setLevel(log_helper.NOTSET)


###############################################################################
# Base Classes
###############################################################################


SUMMARY_INFORMATION_STREAM_NAMES = ('\x05SummaryInformation',
                                    '\x05DocumentSummaryInformation')


class OleRecordFile(olefile.OleFileIO):
    """ an OLE compound file whose streams have (mostly) record structure

    'record structure' meaning that streams are a sequence of records. Records
    are structure with information about type and size in their first bytes
    and type-dependent data of given size after that.

    Subclass of OleFileIO!
    """

    def open(self, filename, *args, **kwargs):
        """Call OleFileIO.open."""
        #super(OleRecordFile, self).open(filename, *args, **kwargs)
        OleFileIO.open(self, filename, *args, **kwargs)

    @classmethod
    def stream_class_for_name(cls, stream_name):
        """ helper for iter_streams, must be overwritten in subclasses

        will not be called for SUMMARY_INFORMATION_STREAM_NAMES
        """
        return OleRecordStream    # this is an abstract class!

    def iter_streams(self):
        """ find all streams, including orphans """
        logger.debug('Finding streams in ole file')

        for sid, direntry in enumerate(self.direntries):
            is_orphan = direntry is None
            if is_orphan:
                # this direntry is not part of the tree --> unused or orphan
                direntry = self._load_direntry(sid)
            is_stream = direntry.entry_type == olefile.STGTY_STREAM
            logger.debug('direntry {:2d} {}: {}'.format(
                sid, '[orphan]' if is_orphan else direntry.name,
                'is stream of size {}'.format(direntry.size) if is_stream else
                'no stream ({})'.format(ENTRY_TYPE2STR[direntry.entry_type])))
            if is_stream:
                if not is_orphan and \
                        direntry.name in SUMMARY_INFORMATION_STREAM_NAMES:
                    clz = OleSummaryInformationStream
                else:
                    clz = self.stream_class_for_name(direntry.name)
                stream = clz(self._open(direntry.isectStart, direntry.size),
                             direntry.size,
                             None if is_orphan else direntry.name,
                             direntry.entry_type)
                yield stream
                stream.close()


class OleRecordStream(object):
    """ a stream found in an OleRecordFile

    Always has a name and a size (both read-only). Has an OleFileStream handle.

    abstract base class
    """

    def __init__(self, stream, size, name, stream_type):
        self.stream = stream
        self.size = size
        self.name = name
        if stream_type not in ENTRY_TYPE2STR:
            raise ValueError('Unknown stream type: {0}'.format(stream_type))
        self.stream_type = stream_type

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
            # logger.debug('Record type {0} of size {1}'
            #              .format(rec_type, rec_size))

            # determine what class to wrap this into
            rec_clz, force_read = self.record_class_for_type(rec_type)

            if fill_data or force_read:
                data = self.stream.read(rec_size)
                if len(data) != rec_size:
                    raise IOError('Unexpected end of stream ({0} < {1})'
                                  .format(len(data), rec_size))
            else:
                self.stream.seek(rec_size, SEEK_CUR)
                data = None
            rec_object = rec_clz(rec_type, rec_size, other, pos, data)

            # "We are microsoft, we do not always adhere to our specifications"
            rec_object.read_some_more(self.stream)
            yield rec_object

    def close(self):
        """Close this stream (i.e. the stream given in constructor)."""
        self.stream.close()

    def __str__(self):
        return '[{0} {1} (type {2}, size {3})' \
               .format(self.__class__.__name__,
                       self.name or '[orphan]',
                       ENTRY_TYPE2STR[self.stream_type],
                       self.size)


class OleSummaryInformationStream(OleRecordStream):
    """ stream for \05SummaryInformation and \05DocumentSummaryInformation

    Do nothing so far. OleFileIO reads quite some info from this. For more info
    see [MS-OSHARED] 2.3.3 and [MS-OLEPS] 2.21 and references therein.

    See also: info read in oleid.py.
    """
    def iter_records(self, fill_data=False):
        """ yields nothing, stops at once """
        return
        yield   # required to make this a generator pylint: disable=unreachable


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
        self.finish_constructing(more_data)

    def finish_constructing(self, more_data):
        """ finish constructing this record

        Can save more_data from OleRecordStream.read_record_head and/or parse
        data (if it was read).

        Base implementation, does nothing. To be overwritten in subclasses.

        Implementations should take into account that self.data may be None.
        Should create the same attributes, whether data is present or not. Eg::

            def finish_constructing(self, more_data):
                self.more = more_data
                self.attr1 = None
                self.attr2 = None
                if self.data:
                    self.attr1, self.attr2 = struct.unpack('<HH', self.data)
        """
        pass

    def read_some_more(self, stream):
        """ Read some more data from stream after end of this record

        Found that for CurrentUserAtom in "Current User" stream of ppt files,
        the last attribute (user name in unicode) is found *behind* the record
        data. Thank you, Microsoft!

        Do this only if you are certain you will not mess up the following
        records!

        This base implementation does nothing. For optional overwriting in
        subclasses (like PptRecordUserAtom where no record should follow.)
        """
        return

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
         must_parse=None, do_per_record=None, verbose=False):
    """ parse all given file names and print rough structure

    if an error occurs while parsing a stream of type in must_parse, the error
    will be raised. Otherwise a message is printed
    """
    log_helper.enable_logging(False, 'debug' if verbose else 'info')
    if do_per_record is None:
        def do_per_record(record):         # pylint: disable=function-redefined
            pass   # do nothing
    if not filenames:
        logger.info('need file name[s]')
        return 2
    for filename in filenames:
        logger.info('checking file {0}'.format(filename))
        if not olefile.isOleFile(filename):
            logger.info('not an ole file - skip')
            continue
        ole = ole_file_class(filename)

        for stream in ole.iter_streams():
            logger.info('  parse ' + str(stream))
            try:
                for record in stream.iter_records():
                    logger.info('    ' + str(record))
                    do_per_record(record)
            except Exception:
                if not must_parse:
                    raise
                elif isinstance(stream, must_parse):
                    raise
                else:
                    logger.info('  failed to parse', exc_info=True)

    log_helper.end_logging()
    return 0


if __name__ == '__main__':
    sys.exit(test(sys.argv[1:]))
