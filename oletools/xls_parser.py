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

from oletools.thirdparty import olefile


# === PYTHON 2+3 SUPPORT ======================================================

if sys.version_info[0] >= 3:
    unichr = chr

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


def read_unicode(data, start_idx, n_chars):
    """ read a unicode string from a XLUnicodeStringNoCch structure """
    # first bit 0x0 --> only low-bytes are saved, all high bytes are 0
    # first bit 0x1 --> 2 bytes per character
    low_bytes_only = (ord(data[start_idx]) == 0)
    if low_bytes_only:
        end_idx = start_idx + 1 + n_chars
        return data[start_idx+1:end_idx].decode('ascii'), end_idx
    else:
        return read_unicode_2byte(data, start_idx+1, n_chars)


def read_unicode_2byte(data, start_idx, n_chars):
    """ read a unicode string with characters encoded by 2 bytes """
    end_idx = start_idx + n_chars * 2
    if n_chars < 256:  # faster version, long format string for unpack
        unichars = (unichr(val) for val in
                    unpack('<' + 'H'*n_chars, data[start_idx:end_idx]))
    else:              # slower version but less memory-extensive
        unichars = (unichr(unpack('<H', data[data_idx:data_idx+2])[0])
                    for data_idx in xrange(start_idx, end_idx, 2))
    return u''.join(unichars), end_idx


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

    stream argument can be oleile.OleStream or ooxml.ZipSubFile
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
        """ iterate over records in streams

        Stream must be positioned at start of records (e.g. start of stream).
        """
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


class XlsbStream(XlsStream):
    """ binary stream of an xlsb file, usually have a record structure """

    HIGH_BIT_MASK = 0b10000000
    LOW7_BIT_MASK = 0b01111111

    def iter_records(self):
        """ iterate over records in stream

        Record type and size are encoded differently than in xls streams.
        (c.f. [MS-XLSB, Paragraph 2.1.4: Record)
        """
        while True:
            pos = self.stream.tell()
            if pos >= self.size:
                break
            val = ord(self.stream.read(1))
            if val & self.HIGH_BIT_MASK:    # high bit of the low byte is 1
                val2 = ord(self.stream.read(1))         # need another byte
                # combine 7 low bits of each byte
                type = (val & self.LOW7_BIT_MASK) + \
                       ((val2 & self.LOW7_BIT_MASK) << 7)
            else:
                type = val

            size = 0
            shift = 0
            for _ in range(4):      # size needs up to 4 byte
                val = ord(self.stream.read(1))
                size += (val & self.LOW7_BIT_MASK) << shift
                shift += 7
                if (val & self.HIGH_BIT_MASK) == 0:   # high-bit is 0 --> done
                    break

            if pos + size > self.size:
                raise ValueError('Stream does not seem to have record '
                                 'structure or is incomplete (record size {0})'
                                 .format(size))
            data = self.stream.read(size)

            clz = XlsbRecord
            if type == XlsbBeginSupBook.TYPE:
                clz = XlsbBeginSupBook
            yield clz(type, size, pos, data)


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

#: records found in xlsb binary parts
FREQUENT_RECORDS_XLSB = dict([
    (360, 'BrtBeginSupBook'),
    (588, 'BrtEndSupBook'),
    (667, 'BrtSupAddin'),
    (355, 'BrtSupBookSrc'),
    (586, 'BrtSupNameBits'),
    (584, 'BrtSupNameBool'),
    (587, 'BrtSupNameEnd'),
    (581, 'BrtSupNameErr'),
    (585, 'BrtSupNameFmla'),
    (583, 'BrtSupNameNil'),
    (580, 'BrtSupNameNum'),
    (582, 'BrtSupNameSt'),
    (577, 'BrtSupNameStart'),
    (579, 'BrtSupNameValueEnd'),
    (578, 'BrtSupNameValueStart'),
    (358, 'BrtSupSame'),
    (357, 'BrtSupSelf'),
    (359, 'BrtSupTabs'),
])


class XlsRecord(object):
    """ basic building block of data in workbook stream """

    #: max size of a record in xls stream (does not apply to xlsb)
    MAX_SIZE = 8224

    # to be overwritten in subclasses that have fixed type/size
    TYPE = None
    SIZE = None

    def __init__(self, type, size, pos, data=None):
        """ create a record """
        self.type = type
        if self.MAX_SIZE is not None and size > self.MAX_SIZE:
            logging.warning('record size {0} exceeds max size'
                            .format(size))
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
        elif self.virt_path == u'\u0020':   # space ; ctab can be anything
            self.support_link_type = self.LINK_TYPE_UNUSED
        elif self.virt_path == u'\u0000':
            self.support_link_type = self.LINK_TYPE_SAMESHEET
        elif self.ctab == 0x0 and self.virt_path:
            self.support_link_type = self.LINK_TYPE_OLE_DDE
        elif self.ctab > 0 and self.virt_path:
            self.support_link_type = self.LINK_TYPE_EXTERNAL

    def _type_str(self):
        return 'SupBook Record ({0})'.format(self.support_link_type)


class XlsbRecord(XlsRecord):
    """ like an xls record, but from binary part of xlsb file

    has no MAX_SIZE and types have different meanings
    """

    MAX_SIZE = None

    def _type_str(self):
        """ simplification for subclasses to create their own __str__ """
        try:
            return FREQUENT_RECORDS_XLSB[self.type]
        except KeyError:
            return 'XlsbRecord type {0}'.format(self.type)


class XlsbBeginSupBook(XlsbRecord):
    """ Record beginning an external link in xlsb file

    contains information about the link itself (e.g. for DDE the link is
    string1 + ' ' + string2)
    """

    TYPE = 360
    LINK_TYPE_WORKBOOK = 'workbook'
    LINK_TYPE_DDE = 'DDE'
    LINK_TYPE_OLE = 'OLE'
    LINK_TYPE_UNEXPECTED = 'unexpected'
    LINK_TYPE_UNKNOWN = 'unknown'

    def __init__(self, *args, **kwargs):
        super(XlsbBeginSupBook, self).__init__(*args, **kwargs)
        self.link_type = self.LINK_TYPE_UNKNOWN
        self.string1 = ''
        self.string2 = ''
        if self.data is None:
            return
        self.sbt = unpack('<H', self.data[0:2])[0]
        if self.sbt == 0:
            self.link_type = self.LINK_TYPE_WORKBOOK
        elif self.sbt == 1:
            self.link_type = self.LINK_TYPE_DDE
        elif self.sbt == 2:
            self.link_type = self.LINK_TYPE_OLE
        else:
            logging.warning('Unexpected link type {0} encountered'
                            .format(self.data[0]))
            self.link_type = self.LINK_TYPE_UNEXPECTED

        start_idx = 2
        n_chars = unpack('<I', self.data[start_idx:start_idx+4])[0]
        if n_chars == 0xFFFFFFFF:
            logging.warning('Max string length 0xFFFFFFF is not allowed')
        elif self.size < n_chars*2 + start_idx+4:
            logging.warning('Impossible string length {0} for data length {1}'
                            .format(n_chars, self.size))
        else:
            self.string1, start_idx = read_unicode_2byte(self.data,
                                                         start_idx+4, n_chars)

        n_chars = unpack('<I', self.data[start_idx:start_idx+4])[0]
        if n_chars == 0xFFFFFFFF:
            logging.warning('Max string length 0xFFFFFFF is not allowed')
        elif self.size < n_chars*2 + start_idx+4:
            logging.warning('Impossible string length {0} for data length {1}'
                            .format(n_chars, self.size) + ' for string2')
        else:
            self.string2, _ = read_unicode_2byte(self.data, start_idx+4,
                                                 n_chars)

    def _type_str(self):
        return 'XlsbBeginSupBook Record ({0}, "{1}", "{2}")' \
               .format(self.link_type, self.string1, self.string2)


###############################################################################
# XLSB Binary Parts
###############################################################################

def parse_xlsb_part(stream, _, filename):
    """ Excel xlsb files also have a record structure. iter records """
    for record in XlsbStream(stream, filename).iter_records():
        yield record


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
