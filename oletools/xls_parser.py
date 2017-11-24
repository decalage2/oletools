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

#------------------------------------------------------------------------------
# CHANGELOG:
# 2017-11-02 v0.01 CH: - first version

__version__ = '0.1'

# -----------------------------------------------------------------------------
#  TODO:
#  - parse more record types (ExternName, ...)
#  - check what bad stuff can be in other storages: Embedded ("MBD..."), Linked
#    ("LNK..."), "MsoDataStore" and OleStream ('\001Ole')
#
# -----------------------------------------------------------------------------
#  REFERENCES:
#  - [MS-XLS]: Excel Binary File Format (.xls) Structure Specification
#    https://msdn.microsoft.com/en-us/library/office/cc313154(v=office.14).aspx
#  - Understanding the Excel .xls Binary File Format
#    https://msdn.microsoft.com/en-us/library/office/gg615597(v=office.14).aspx
#
# -- IMPORTS ------------------------------------------------------------------

import sys
import os.path
from struct import unpack
from io import SEEK_CUR
import logging

# little hack to allow absolute imports even if oletools is not installed.
# Copied from olevba.py
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))  # pylint: disable=invalid-name
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))             # pylint: disable=invalid-name
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty import olefile                                # pylint: disable=wrong-import-position


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


def is_xls(filename):
    """
    determine whether a given file is an excel ole file

    returns True if given file is an ole file and contains a Workbook stream

    todo: could further check that workbook stream starts with a globals
    substream
    """
    try:
        for stream in XlsFile(filename).get_streams():
            if isinstance(stream, WorkbookStream):
                return True
    except Exception:
        return False


###############################################################################
# File, Storage, Stream
###############################################################################


class XlsFile(olefile.OleFileIO):
    """ specialization of an OLE compound file """

    def get_streams(self):
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
                if direntry.name == 'Workbook':
                    clz = WorkbookStream
                else:
                    clz = XlsStream
                yield clz(self._open(direntry.isectStart, direntry.size),
                          None if is_orphan else direntry.name)


class XlsStream(object):
    """ specialization of an OLE stream

    Currently not much use, but may be interesting for further sub-classing
    when extending this code.
    """

    def __init__(self, stream, name):
        self.stream = stream
        self.size = stream.size
        self.name = name

    def __str__(self):
        return '[XlsStream {0} (size {1})' \
               .format(self.name or '[orphan]', self.size)


class WorkbookStream(XlsStream):
    """ the workbook stream which contains records """

    def iter_records(self, fill_data=False):
        """ iterate over records in streams"""
        if self.stream.tell() != 0:
            logging.debug('have to jump to start')
            self.stream.seek(0)

        while True:
            # unpacking as in olevba._extract_vba
            pos = self.stream.tell()
            if pos >= self.size:
                break
            type = unpack('<H', self.stream.read(2))[0]
            size = unpack('<H', self.stream.read(2))[0]
            force_read = False
            if type == XlsRecordBof.TYPE:
                clz = XlsRecordBof
                force_read = True
            elif type == XlsRecordEof.TYPE:
                clz = XlsRecordEof
            elif type == XlsRecordSupBook.TYPE:
                clz = XlsRecordSupBook
                force_read = True
            else:
                clz = XlsRecord
            data = None
            if fill_data or force_read:
                data = self.stream.read(size)
            else:
                self.stream.seek(size, SEEK_CUR)
            yield clz(type, size, pos, data)

    def __str__(self):
        return '[Workbook Stream (size {0})'.format(self.size)


###############################################################################
# RECORDS
###############################################################################

# records that appear often but do not need their own XlsRecord subclass (yet)
FREQUENT_RECORDS = dict([
    ( 156, 'BuiltInFnGroupCount'),             # pylint: disable=bad-whitespace
    (2147, 'BookExt'),                         # pylint: disable=bad-whitespace
    ( 442, 'CodeName'),                        # pylint: disable=bad-whitespace
    (  66, 'CodePage'),                        # pylint: disable=bad-whitespace
    (4195, 'Dat'),                             # pylint: disable=bad-whitespace
    (2154, 'DataLabExt'),                      # pylint: disable=bad-whitespace
    (2155, 'DataLabExtContents'),              # pylint: disable=bad-whitespace
    ( 215, 'DBCell'),                          # pylint: disable=bad-whitespace
    ( 220, 'DbOrParmQry'),                     # pylint: disable=bad-whitespace
    (2051, 'DBQueryExt'),                      # pylint: disable=bad-whitespace
    (2166, 'DConn'),                           # pylint: disable=bad-whitespace
    (  35, 'ExternName'),                      # pylint: disable=bad-whitespace
    (  23, 'ExternSheet'),                     # pylint: disable=bad-whitespace
    ( 255, 'ExtSST'),                          # pylint: disable=bad-whitespace
    (2052, 'ExtString'),                       # pylint: disable=bad-whitespace
    (2151, 'FeatHdr'),                         # pylint: disable=bad-whitespace
    (  91, 'FileSharing'),                     # pylint: disable=bad-whitespace
    (1054, 'Format'),                          # pylint: disable=bad-whitespace
    (  49, 'Font'),                            # pylint: disable=bad-whitespace
    (2199, 'GUIDTypeLib'),                     # pylint: disable=bad-whitespace
    ( 440, 'HLink'),                           # pylint: disable=bad-whitespace
    ( 225, 'InterfaceHdr'),                    # pylint: disable=bad-whitespace
    ( 226, 'InterfaceEnd'),                    # pylint: disable=bad-whitespace
    ( 523, 'Index'),                           # pylint: disable=bad-whitespace
    (  24, 'Lbl'),                             # pylint: disable=bad-whitespace
    ( 193, 'Mms'),                             # pylint: disable=bad-whitespace
    (  93, 'Obj'),                             # pylint: disable=bad-whitespace
    (4135, 'ObjectLink'),                      # pylint: disable=bad-whitespace
    (2058, 'OleDbConn'),                       # pylint: disable=bad-whitespace
    ( 222, 'OleObjectSize'),                   # pylint: disable=bad-whitespace
    (2214, 'RichTextStream'),                  # pylint: disable=bad-whitespace
    (2146, 'SheetExt'),                        # pylint: disable=bad-whitespace
    (1212, 'ShrFmla'),                         # pylint: disable=bad-whitespace
    (2060, 'SxViewExt'),                       # pylint: disable=bad-whitespace
    (2136, 'SxViewLink'),                      # pylint: disable=bad-whitespace
    (2049, 'WebPub'),                          # pylint: disable=bad-whitespace
    ( 224, 'XF (formatting)'),                 # pylint: disable=bad-whitespace
    (2173, 'XFExt (formatting)'),              # pylint: disable=bad-whitespace
    ( 659, 'Style'),                           # pylint: disable=bad-whitespace
    (2194, 'StyleExt')                         # pylint: disable=bad-whitespace
])


class XlsRecord(object):
    """ basic building block of data in workbook stream """

    #: max size of a record
    MAX_SIZE = 8224

    # to be overwritten in subclasses that have fixed type/size
    TYPE = None
    SIZE = None

    def __init__(self, type, size, pos, data=None):
        """ create a record """
        self.type = type
        if size > self.MAX_SIZE:
            raise ValueError('size {0} exceeds max size'.format(size))
        elif self.SIZE is not None and size != self.SIZE:
            raise ValueError('size {0} is not as expected for this type'
                             .format(size))
        self.size = size
        self.pos = pos
        self.data = data
        if data is not None and len(data) != size:
            raise ValueError('data size {0} is not expected size {1}'
                             .format(len(data), size))

    def read_data(self, stream):
        """ read data from stream if up to now only pos was known """
        raise NotImplementedError()

    def _type_str(self):
        """ simplification for subclasses to create their own __str__ """
        try:
            return FREQUENT_RECORDS[self.type]
        except KeyError:
            return 'XlsRecord type {0}'.format(self.type)

    def __str__(self):
        return '[' + self._type_str() + \
               ' (size {0} from {1})]'.format(self.size, self.pos)


class XlsRecordBof(XlsRecord):
    """ record found at beginning of substreams """
    TYPE = 2057
    SIZE = 16
    # types of substreams
    DOCTYPES = dict([(0x5, 'workbook'), (0x10, 'dialog/worksheet'),
                     (0x20, 'chart'), (0x40, 'macro')])

    def __init__(self, *args, **kwargs):
        super(XlsRecordBof, self).__init__(*args, **kwargs)
        if self.data is None:
            self.doctype = None
            return
        # parse data (only doctype, ignore rest)
        self.doctype = unpack('<H', self.data[2:4])[0]

    def _type_str(self):
        return 'BOF Record ({0} substream)'.format(
            self.DOCTYPES[self.doctype] if self.doctype in self.DOCTYPES
            else 'unknown')


class XlsRecordEof(XlsRecord):
    """ record found at end of substreams """
    TYPE = 10
    SIZE = 0

    def _type_str(self):
        return 'EOF Record'


class XlsRecordSupBook(XlsRecord):
    """ The SupBook record specifies a supporting link

    "... The collection of records specifies the contents of an external
    workbook, DDE data source, or OLE data source." (MS-XLS, paragraph 2.4.271)
    """

    TYPE = 430

    LINK_TYPE_UNKNOWN = 'unknown'
    LINK_TYPE_SELF = 'self-referencing'
    LINK_TYPE_ADDIN = 'addin-referencing'
    LINK_TYPE_UNUSED = 'unused'
    LINK_TYPE_SAMESHEET = 'same-sheet'
    LINK_TYPE_OLE_DDE = 'ole/dde data source'
    LINK_TYPE_EXTERNAL = 'external workbook'

    def __init__(self, *args, **kwargs):
        super(XlsRecordSupBook, self).__init__(*args, **kwargs)

        # set defaults
        self.ctab = None
        self.cch = None
        self.virt_path = None
        self.support_link_type = self.LINK_TYPE_UNKNOWN
        if self.data is None:
            return

        # parse data
        if self.size < 4:
            raise ValueError('not enough data (size is {0} but need >= 4)'
                             .format(self.size))
        self.ctab, self.cch = unpack('<HH', self.data[:4])
        if 0 < self.cch <= 0xff:
            # this is the length of virt_path
            self.virt_path, _ = read_unicode(self.data, 4, self.cch)
        else:
            self.virt_path, _ = u'', 4
        # ignore variable rgst

        if self.cch == 0x401:    # ctab is undefined and to be ignored
            self.support_link_type = self.LINK_TYPE_SELF
        elif self.ctab == 0x1 and self.cch == 0x3A01:
            self.support_link_type = self.LINK_TYPE_ADDIN
            # next records must be ExternName with all add-in functions
        elif self.virt_path == '\u0020':   # space ; ctab can be anything
            self.support_link_type = self.LINK_TYPE_UNUSED
        elif self.virt_path == '\u0000':
            self.support_link_type = self.LINK_TYPE_SAMESHEET
        elif self.ctab == 0x0 and self.virt_path:
            self.support_link_type = self.LINK_TYPE_OLE_DDE
        elif self.ctab > 0 and self.virt_path:
            self.support_link_type = self.LINK_TYPE_EXTERNAL

    def _type_str(self):
        return 'SupBook Record ({0})'.format(self.support_link_type)


def read_unicode(data, start_idx, n_chars):
    """ read a unicode string from a XLUnicodeStringNoCch structure """
    # first bit 0x0 --> only low-bytes are saved, all high bytes are 0
    # first bit 0x1 --> 2 bytes per character
    low_bytes_only = (ord(data[start_idx]) == 0)
    if low_bytes_only:
        end_idx = start_idx + 1 + n_chars
        return data[start_idx+1:end_idx].decode('ascii'), end_idx
    end_idx = start_idx + 1 + n_chars * 2
    return u''.join(unichr(val) for val in
                    unpack('<' + 'H'*n_chars, data[start_idx+1:end_idx])), \
           end_idx


###############################################################################
# TESTING
###############################################################################


def test(*filenames):
    """ parse all given file names and print rough structure """
    logging.basicConfig(level=logging.DEBUG)
    if not filenames:
        logging.info('need file name[s]')
        return 2
    for filename in filenames:
        logging.info('checking file {0}'.format(filename))
        if not olefile.isOleFile(filename):
            logging.info('not an ole file - skip')
            continue
        xls = XlsFile(filename)

        for stream in xls.get_streams():
            logging.info(stream)
            if isinstance(stream, WorkbookStream):
                for record in stream.iter_records():
                    logging.info('  {0}'.format(record))
    return 0


if __name__ == '__main__':
    sys.exit(test(*sys.argv[1:]))
