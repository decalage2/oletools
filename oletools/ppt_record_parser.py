#!/usr/bin/env python

"""
ppt_record_parser.py

Alternative to ppt_parser.py that works on records
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

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2017-11-30 v0.01 CH: - first version, can be used in oledump

# -----------------------------------------------------------------------------
# TODO:
# - provide stuff from ppt_parser as well and replace it

# -----------------------------------------------------------------------------
#  REFERENCES:
#  - [MS-PPT]


import sys
from struct import unpack      # unsigned: 1 Byte = B, 2 Byte = H, 4 Byte = L
import logging
import io
import zlib

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
try:
    from oletools import record_base
except ImportError:
    import os.path
    PARENT_DIR = os.path.normpath(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    if PARENT_DIR not in sys.path:
        sys.path.insert(0, PARENT_DIR)
    del PARENT_DIR
    from oletools import record_base


# types of relevant records (there are much more than listed here)
RECORD_TYPES = dict([
    # file structure types
    (0x0ff5, 'UserEditAtom'),
    (0x0ff6, 'CurrentUserAtom'),        # --> use PptRecordCurrentUser instead
    (0x1772, 'PersistDirectoryAtom'),
    (0x2f14, 'CryptSession10Container'),
    # document types
    (0x03e8, 'DocumentContainer'),
    (0x0fc9, 'HandoutContainer'),
    (0x03f0, 'NotesContainer'),
    (0x03ff, 'VbaInfoContainer'),
    (0x03e9, 'DocumentAtom'),
    (0x03ea, 'EndDocumentAtom'),
    # slide types
    (0x03ee, 'SlideContainer'),
    (0x03f8, 'MainMasterContainer'),
    # external object ty
    (0x0409, 'ExObjListContainer'),
    (0x1011, 'ExOleVbaActiveXAtom'),    # --> use PptRecordExOleVbaActiveXAtom
    (0x1006, 'ExAviMovieContainer'),
    (0x100e, 'ExCDAudioContainer'),
    (0x0fee, 'ExControlContainer'),
    (0x0fd7, 'ExHyperlinkContainer'),
    (0x1007, 'ExMCIMovieContainer'),
    (0x100d, 'ExMIDIAudioContainer'),
    (0x0fcc, 'ExOleEmbedContainer'),
    (0x0fce, 'ExOleLinkContainer'),
    (0x100f, 'ExWAVAudioEmbeddedContainer'),
    (0x1010, 'ExWAVAudioLinkContainer'),
    (0x1004, 'ExMediaAtom'),
    (0x040a, 'ExObjListAtom'),
    (0x0fcd, 'ExOleEmbedAtom'),
    (0x0fc3, 'ExOleObjAtom'),           # --> use PptRecordExOleObjAtom instead
    # other types
    (0x0fc1, 'MetafileBlob'),
    (0x0fb8, 'FontEmbedDataBlob'),
    (0x07e7, 'SoundDataBlob'),
    (0x138b, 'BinaryTagDataBlob'),
    (0x0fba, 'CString'),
])


# record types where version is not 0x0 or 0x1 or 0xf
VERSION_EXCEPTIONS = dict([
    (0x0400, 2),                       # rt_vbainfoatom
    (0x03ef, 2),                       # rt_slideatom
    (0xe9c7, 7),    # tests/test-data/encrypted/encrypted.ppt, not investigated
])


# record types where instance is not 0x0 or 0x1
INSTANCE_EXCEPTIONS = dict([
    (0x0fba, (2, 0x14)),                 # rt_cstring,
    (0x0ff0, (2, 2)),                    # rt_slidelistwithtext,
    (0x0fd9, (3, 4)),                    # rt_headersfooters,
    (0x07e4, (5, 5)),                    # rt_soundcollection,
    (0x03fb, (7, 7)),                    # rt_guideatom,
    (0x07e9, (2, 2)),                    # rt_bookmarkseeatom,
    (0x07f0, (6, 6)),                    # rt_colorschemeatom,
    (0xf125, (0, 5)),                    # rt_timeconditioncontainer,
    (0xf13d, (0, 0xa)),                  # rt_timepropertylist,
    (0x0fc8, (2, 2)),                    # rt_kinsoku,
    (0x0fd2, (3, 3)),                    # rt_kinsokuatom,
    (0x0f9f, (0, 5)),                    # rt_textheaderatom,
    (0x0fb7, (0, 128)),                  # rt_fontentityatom,
    (0x0fa3, (0, 8)),                    # rt_textmasterstyleatom,
    (0x0fad, (0, 8)),                    # rt_textmasterstyle9atom,
    (0x0fb2, (0, 8)),                    # rt_textmasterstyle10atom,
    (0x07f9, (0, 0x80)),                 # rt_blibentitiy9atom,
    (0x0faf, (0, 5)),                    # rt_outlinetextpropsheader9atom,
    (0x0fb8, (0, 3)),                    # rt_fontembeddatablob,
])


def is_ppt(filename):
    """ determine whether given file is a PowerPoint 2003 (ppt) OLE file

    Tries to ppt-parse the file, return False if that fails. Looks for certain
    required streams and records.

    Param filename can be anything that OleFileIO constructor accepts: name of
    file or file data or data stream.

    Will not try to decrypt the file not even try to determine whether it is
    encrypted. If the file is encrypted will either raise an error or just
    return `False`.

    see also: oleid.OleID.check_powerpoint
    """
    have_current_user = False
    have_user_edit = False
    have_persist_dir = False
    have_document_container = False
    ppt_file = None
    try:
        ppt_file = PptFile(filename)
        for stream in ppt_file.iter_streams():
            if stream.name == 'Current User':
                for record in stream.iter_records():
                    if isinstance(record, PptRecordCurrentUser):
                        have_current_user = True
                        if have_current_user and have_user_edit and \
                                have_persist_dir and have_document_container:
                            return True
            elif stream.name == 'PowerPoint Document':
                for record in stream.iter_records():
                    if record.type == 0x0ff5:     # UserEditAtom
                        have_user_edit = True
                    elif record.type == 0x1772:   # PersistDirectoryAtom
                        have_persist_dir = True
                    elif record.type == 0x03e8:   # DocumentContainer
                        have_document_container = True
                    else:
                        continue
                    if have_current_user and have_user_edit and \
                            have_persist_dir and have_document_container:
                        return True
            else:   # ignore other streams/storages since they are optional
                continue
    except Exception as exc:
        logging.debug('Ignoring exception in is_ppt, assume is not ppt',
                      exc_info=True)
    finally:
        if ppt_file is not None:
            ppt_file.close()
    return False


class PptFile(record_base.OleRecordFile):
    """ Record-based view on a PowerPoint ppt file

    This is a subclass of OleFileIO, so can be constructed from file name or
    file data or data stream.
    """

    @classmethod
    def stream_class_for_name(cls, stream_name):
        return PptStream


class PptStream(record_base.OleRecordStream):
    """ a stream of records in a ppt file """

    def read_record_head(self):
        """ read first few bytes of record to determine size and type

        returns (type, size, other) where other is (instance, version)
        """
        ver_inst, rec_type, rec_size = unpack('<HHL', self.stream.read(8))
        instance, version = divmod(ver_inst, 2**4)
        return rec_type, rec_size, (instance, version)

    @classmethod
    def record_class_for_type(cls, rec_type):
        """ determine a class for given record type

        returns (clz, force_read)
        """
        if rec_type == PptRecordCurrentUser.TYPE:
            return PptRecordCurrentUser, True
        elif rec_type == PptRecordExOleObjAtom.TYPE:
            return PptRecordExOleObjAtom, True
        elif rec_type == PptRecordExOleVbaActiveXAtom.TYPE:
            return PptRecordExOleVbaActiveXAtom, True

        try:
            record_name = RECORD_TYPES[rec_type]
            if record_name.endswith('Container'):
                is_container = True
            elif record_name.endswith('Atom'):
                is_container = False
            elif record_name.endswith('Blob'):
                is_container = False
            elif record_name == 'CString':
                is_container = False
            else:
                logging.warning('Unexpected name for record type "{0}". typo?'
                                .format(record_name))
                is_container = False

            if is_container:
                return PptContainerRecord, True
            else:
                return PptRecord, False
        except KeyError:
            return PptRecord, False


class PptRecord(record_base.OleRecordBase):
    """ A Record within a ppt file; has instance and version fields """

    # fixed values for instance and version (usually ver is 0 or 0xf, inst 0/1)
    INSTANCE = None
    VERSION = None

    def finish_constructing(self, more_data):
        """ check and save instance and version """
        instance, version = more_data
        if self.INSTANCE is not None and self.INSTANCE != instance:
            raise ValueError('invalid instance {0} for {1}'
                             .format(instance, self))
        elif self.INSTANCE is not None and instance not in (0, 1):
            try:
                min_val, max_val = INSTANCE_EXCEPTIONS[self.type]
                is_ok = (min_val <= instance <= max_val)
            except KeyError:
                is_ok = False
            if not is_ok:
                logging.warning('unexpected instance {0} for {1}'
                                .format(instance, self))
        self.instance = instance
        if self.VERSION is not None and self.VERSION != version:
            raise ValueError('invalid version {0} for {1}'
                             .format(version, self))
        elif self.VERSION is None and version not in (0x0, 0x1, 0xf):
            try:
                is_ok = version == VERSION_EXCEPTIONS[self.type]
            except KeyError:
                is_ok = False
            if not is_ok:
                logging.warning('unexpected version {0} for {1}'
                                .format(version, self))
        self.version = version

    def _type_str(self):
        """ helper for __str__, base implementation """
        try:
            record_name = RECORD_TYPES[self.type]
            return '{0} record'.format(record_name)
        except KeyError:
            return '{0} type 0x{1:04x}'.format(self.__class__.__name__,
                                               self.type)


class PptContainerRecord(PptRecord):
    """ A record that contains other records """

    def finish_constructing(self, more_data):
        """ parse records from self.data """
        # set self.version and self.instance
        super(PptContainerRecord, self).finish_constructing(more_data)
        self.records = None
        if not self.data:
            return

        # logging.debug('parsing contents of container record {0}'
        #               .format(self))

        # create a stream from self.data and parse it like any other
        data_stream = io.BytesIO(self.data)
        record_stream = PptStream(data_stream, self.size,
                                  'PptContainerRecordSubstream',
                                  record_base.STGTY_SUBSTREAM)
        self.records = list(record_stream.iter_records())
        # logging.debug('done parsing contents of container record {0}'
        #               .format(self))

    def __str__(self):
        text = super(PptContainerRecord, self).__str__()
        if self.records is None:
            return '{0}, unparsed{1}'.format(text[:-2], text[-2:])
        elif self.records:
            return '{0}, contains {1} recs{2}' \
                   .format(text[:-2], len(self.records), text[-2:])
        else:
            return text


class PptRecordCurrentUser(PptRecord):
    """ The CurrentUserAtom record """
    TYPE = 0x0ff6
    VERSION = 0
    INSTANCE = 0

    def finish_constructing(self, more_data):
        """ read various attributes from data """
        super(PptRecordCurrentUser, self).finish_constructing(more_data)
        if self.size < 24:
            raise ValueError('CurrentUser record is too small ({0})'
                             .format(self.size))
        self.size2 = None
        self.header_token = None
        self.offset_to_current_edit = None
        self.len_user_name = None
        self.doc_file_version = None
        self.major_version = None
        self.minor_version = None
        self.ansi_user_name = None
        self.unicode_user_name = None

        if not self.data:
            return

        self.size2, self.header_token, self.offset_to_current_edit, \
            self.len_user_name, self.doc_file_version, self.major_version, \
            self.minor_version, _ = unpack('<LLLHHBBH', self.data[0:20])
        if self.size2 != 0x14:
            raise ValueError('Wrong size2 ({0}) in CurrentUser record'
                             .format(self.size2))
        elif self.header_token not in (0xE391C05F, 0xF3D1C4DF):
            raise ValueError('Wrong header_token ({0}) in CurrentUser record'
                             .format(self.header_token))
        elif self.doc_file_version != 0x03F4:
            raise ValueError('Wrong doc file version ({0}) in CurrentUser '
                             'record'.format(self.doc_file_version))
        elif self.major_version != 0x03:
            raise ValueError('Wrong major version ({0}) in CurrentUser record'
                             .format(self.major_version))
        elif self.minor_version != 0x00:
            raise ValueError('Wrong minor version ({0}) in CurrentUser record'
                             .format(self.minor_version))
        self.ansi_user_name = self.data[20:20+self.len_user_name]
        if len(self.ansi_user_name) != self.len_user_name:
            raise ValueError('CurrentUser record is too small for user name '
                             '({0} != {1})'.format(len(self.ansi_user_name),
                                                   self.len_user_name))
        offset = 20 + self.len_user_name
        self.release_version = unpack('<L', self.data[offset:offset+4])[0]
        if self.release_version not in (8, 9):
            raise ValueError('CurrentUser record has wrong release version {0}'
                             .format(self.release_version))
        offset += 4
        if self.size == offset:
            self.unicode_user_name = None    # may be omitted
        elif self.size == offset + 2*self.len_user_name:
            self.unicode_user_name = self.data[offset:].decode('utf-16')
        else:
            raise ValueError('CurrentUser record has wrong size ({0} left)'
                             .format(self.size - offset))

    def is_document_encrypted(self):
        """determine from header_token whether document stream is encrypted"""
        if self.header_token is None:
            raise ValueError('unknown')
        return self.header_token == 0xF3D1C4DF

    def read_some_more(self, stream):
        """ check if unicode user name comes in stream after record

        Can safely do this since no data should come after this record.
        """
        more_data = stream.read(3*self.len_user_name)   # limit data to read
        if self.unicode_user_name is None and \
                len(more_data) == 2*self.len_user_name:
            self.unicode_user_name = more_data.decode('utf-16')
            logging.debug('found unicode user name BEHIND current user atom')
        else:
            logging.warning('Unexplained data of size {0} in "Current User" '
                            'stream'.format(len(more_data)))


class PptRecordExOleObjAtom(PptRecord):
    """ Record that contains info about type of embedded object """

    TYPE = 0x0fc3

    OBJ_TYPES = dict([(0, 'embedded'), (1, 'link'), (2, 'ActiveX')])
    SUB_TYPES = dict([
        (0x00, 'default'),
        (0x01, 'clipart'),
        (0x02, 'word doc'),
        (0x03, 'excel sheet'),
        (0x04, 'MS graph'),
        (0x05, 'MS org chart'),
        (0x06, 'equation'),
        (0x07, 'word art'),
        (0x08, 'sound'),
        (0x0c, 'MS project'),
        (0x0d, 'note-it'),
        (0x0e, 'excel chart'),
        (0x0f, 'media'),
        (0x10, 'WordPad doc'),
        (0x11, 'visio drawing'),
        (0x12, 'OpenDoc text'),
        (0x13, 'OpenDoc calc'),
        (0x14, 'OpenDoc present'),
    ])

    def finish_constructing(self, more_data):
        """ parse some more data from this """
        self.draw_aspect = None
        self.obj_type = None
        self.ex_obj_id = None
        self.sub_type = None
        self.persist_id_ref = None
        if self.size != 0x18:
            raise ValueError('ExOleObjAtom has wrong size {0} != 0x18'
                             .format(self.size))
        if self.data:
            self.draw_aspect, self.obj_type, self.ex_obj_id, self.sub_type, \
                self.persist_id_ref, _ = unpack('<LLLLLL', self.data)
            if self.obj_type not in self.OBJ_TYPES:
                logging.warning('Unknown "type" value in ExOleObjAtom: {0}'
                                .format(self.obj_type))
            if self.sub_type not in self.SUB_TYPES:
                logging.warning('Unknown sub type value in ExOleObjAtom: {0}'
                                .format(self.sub_type))

    def _type_str(self):
        return 'ExOleObjAtom type {0}/{1}'.format(
            self.OBJ_TYPES.get(self.obj_type, str(self.obj_type)),
            self.SUB_TYPES.get(self.sub_type, str(self.sub_type)))


class IterStream(io.RawIOBase):
    """ make a read-only, seekable bytes-stream from an iterable

    Copied from stackoverflow answer by Mechanical snail from Nov 18th 2013
    https://stackoverflow.com/a/20260030/4405656 and extended.

    See also (and maybe could some day merge with): ooxml.ZipSubFile;
    also: oleobj.FakeFile
    """

    def __init__(self, iterable_creator, size=None):
        """ create a Stream using a function that creates the iterable """
        super(IterStream, self).__init__()
        self.iterable_creator = iterable_creator
        self.size = size
        # logging.debug('IterStream.size is {0}'.format(self.size))
        self.reset()

    def reset(self):
        """ re-set array to state right after creation """
        self.iterable = None
        self.leftover = None
        self.at_end = False
        self.curr_pos = 0

    def writable(self):
        return False

    def readable(self):
        return True

    def seekable(self):
        return True

    def readinto(self, target):
        """ read as much data from iterable as necessary to fill target """
        # logging.debug('IterStream.readinto size {0}'.format(len(target)))
        if self.at_end:
            # logging.debug('IterStream: we are at (fake) end')
            return 0
        if self.iterable is None:
            self.iterable = self.iterable_creator()
            # logging.debug('IterStream: created iterable {0}'
            #               .format(self.iterable))
            self.curr_pos = 0
        try:
            target_len = len(target)  # we should return at most this much
            chunk = self.leftover or next(self.iterable)
            # logging.debug('IterStream: chunk is size {0}'.format(len(chunk)))
            output, self.leftover = chunk[:target_len], chunk[target_len:]
            # logging.debug('IterStream: output is size {0}, leftover is {1}'
            #               .format(len(output), len(self.leftover)))
            target[:len(output)] = output
            self.curr_pos += len(output)
            # logging.debug('IterStream: pos updated to {0}'
            #               .format(self.curr_pos))
            return len(output)
        except StopIteration:
            # logging.debug('IterStream: source iterable exhausted')
            self.at_end = True
            return 0    # indicate EOF

    def seek(self, offset, whence=io.SEEK_SET):
        """ can seek to start, possibly end """
        if offset != 0 and whence == io.SEEK_SET:
            # logging.debug('IterStream: trying to seek to offset {0}.'
            #               .format(offset))
            if offset > self.curr_pos:
                self.readinto(bytearray(offset - self.curr_pos))
            elif offset == self.curr_pos:
                pass
            else:   # need to re-create iterable
                self.reset()
                self.readinto(bytearray(offset))
            if self.curr_pos != offset:
                # logging.debug('IterStream: curr_pos {0} != offset {1}!'
                #               .format(self.curr_pos, offset))
                raise RuntimeError('programming error in IterStream.tell!')
            return self.curr_pos
        elif whence == io.SEEK_END:  # seek to end
            # logging.debug('IterStream: seek to end')
            if self.size is None:
                # logging.debug('IterStream: trying to seek to end but size '
                #               'unknown --> raise IOError')
                raise IOError('size unknown, cannot seek to end')
            self.at_end = True   # fake jumping to the end
            self.iterable = None   # cannot safely be used any more
            self.leftover = None
            return self.size
        elif whence == io.SEEK_SET:   # seek to start
            # logging.debug('IterStream: seek to start')
            self.reset()
            return 0
        elif whence == io.SEEK_CUR:   # e.g. called by tell()
            # logging.debug('IterStream: seek to curr pos')
            if self.at_end:
                return self.size
            return self.curr_pos
        elif whence not in (io.SEEK_SET, io.SEEK_CUR, io.SEEK_END):
            # logging.debug('Illegal 2nd argument to seek(): {0}'
            #               .format(whence))
            raise IOError('Illegal 2nd argument to seek(): {0}'.format(whence))
        else:
            # logging.debug('not implemented: {0}, {1}'.format(offset, whence))
            raise NotImplementedError('seek only partially implemented. '
                                      'Cannot yet seek to {0} from {1}'
                                      .format(offset, whence))

    def close(self):
        self.iterable = None
        self.leftover = None
        self.at_end = False
        self.curr_pos = 0


class PptRecordExOleVbaActiveXAtom(PptRecord):
    """ record that contains and ole object / vba storage / active x control

    Contains the actual data of the ole object / VBA storage / ActiveX control
    in compressed or uncompressed form.

    Corresponding types in [MS-PPT]:
    ExOleObjStg, ExOleObjStgUncompressedAtom, ExOleObjStgCompressedAtom,
    VbaProjectStg, VbaProjectStgUncompressedAtom, VbaProjectStgCompressedAtom,
    ExControlStg, ExControlStgUncompressedAtom, ExControlStgCompressedAtom.

    self.data is "An array of bytes that specifies a structured storage
    (described in [MSDN-COM]) for the OLE object / ActiveX control / VBA
    project ([MS-OVBA] section 2.2.1)."
    If compressed, "The original bytes of the storage are compressed by the
    algorithm specified in [RFC1950] and are decompressed by the algorithm
    specified in [RFC1951]."   (--> meaning zlib)
    "Office Forms ActiveX controls are specified in [MS-OFORMS]."

    whether this is an OLE object or ActiveX control or a VBA Storage, need to
    find the corresponding PptRecordExOleObjAtom
    TODO: do that!
    """

    TYPE = 0x1011

    def is_compressed(self):
        """ determine whether data is compressed or uncompressed """
        return self.instance == 1

    def get_uncompressed_size(self):
        """ Get size of data in uncompressed form

        For uncompressed data, this just returns self.size. For compressed
        data, this reads and returns the doecmpressedSize field value from
        self.data.  Raises a value error if compressed and data is not
        available.
        """
        if not self.is_compressed():
            return self.size
        elif self.data is None:
            raise ValueError('Data not read from record')
        else:
            return unpack('<L', self.data[:4])[0]

    def iter_uncompressed(self, chunk_size=4096):
        """ iterate over data, decompress data if necessary

        chunk_size is used for input to decompression, so chunks yielded from
        this may well be larger than that. Last chunk is most probably smaller.
        """
        if self.data is None:
            raise ValueError('data not read from record')
        must_decomp = self.is_compressed()
        start_idx = 0
        out_size = 0
        if must_decomp:
            decompressor = zlib.decompressobj()
            start_idx = 4
        while start_idx < self.size:
            end_idx = min(self.size, start_idx+chunk_size)
            if must_decomp:
                result = decompressor.decompress(decompressor.unconsumed_tail +
                                                 self.data[start_idx:end_idx])
            else:
                result = self.data[start_idx:end_idx]
            yield result
            logging.debug('decompressing from {0} to {1} resulted in {2} new'
                          .format(start_idx, end_idx, len(result)))
            out_size += len(result)
            start_idx = end_idx
        if must_decomp:
            result = decompressor.flush()
            out_size += len(result)
            yield result
        if out_size != self.get_uncompressed_size():
            logging.warning('Decompressed data has wrong size {0} != {1}'
                            .format(out_size, self.get_uncompressed_size()))

    def get_data_as_olefile(self, debug_output=False):
        """ return an OleFileIO that streams from iter_uncompressed

        probably only works if data is an OLE object, otherwise expect
        exception
        """
        if debug_output:
            record_base.enable_olefile_logging()
        return record_base.OleFileIO(IterStream(self.iter_uncompressed,
                                                self.get_uncompressed_size()),
                                     debug=debug_output)

    def __str__(self):
        text = super(PptRecordExOleVbaActiveXAtom, self).__str__()
        compr_text = 'compressed' if self.is_compressed() else 'uncompressed'
        return '{0}, {1}{2}'.format(text[:-2], compr_text, text[-2:])


###############################################################################
# TESTING
###############################################################################


def print_records(record, print_fn, indent, do_print_record):
    """ print additional info for record

    prints additional info for some types and subrecords recursively
    """
    if do_print_record:
        print_fn('{0}{1}'.format('  ' * indent, record))
    if isinstance(record, PptContainerRecord):
        for subrec in record.records:
            print_records(subrec, print_fn, indent+1, True)
    elif isinstance(record, PptRecordCurrentUser):
        logging.info('{4}--> crypt: {0}, offset {1}, user {2}/{3}'
                     .format(record.is_document_encrypted(),
                             record.offset_to_current_edit,
                             repr(record.ansi_user_name),
                             repr(record.unicode_user_name),
                             '  ' * indent))
    elif isinstance(record, PptRecordExOleObjAtom):
        logging.info('{2}--> obj id {0}, persist id ref {1}'
                     .format(record.ex_obj_id, record.persist_id_ref,
                             '  ' * indent))
    elif isinstance(record, PptRecordExOleVbaActiveXAtom):
        ole = record.get_data_as_olefile()
        for entry in ole.listdir():
            logging.info('{0}ole entry {1}'.format('  ' * indent, entry))


if __name__ == '__main__':
    def do_per_record(record):
        print_records(record, logging.info, 2, False)
    sys.exit(record_base.test(sys.argv[1:], PptFile,
                              do_per_record=do_per_record,
                              verbose=False))
