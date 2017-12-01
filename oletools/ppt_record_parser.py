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

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2017-11-30 v0.01 CH: - first version based on xls_parser

#------------------------------------------------------------------------------
# TODO:

# -----------------------------------------------------------------------------
#  REFERENCES:
#  - [MS-PPT]


import sys
from struct import unpack
import logging
import record_base
import io


class PptFile(record_base.OleRecordFile):
    """ Record-based view on a PowerPoint ppt file """

    @classmethod
    def stream_class_for_name(self, stream_name):
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

        try:
            record_name = RECORD_TYPES[rec_type]
            if record_name.endswith('Container'):
                is_container = True
            elif record_name.endswith('Atom'):
                is_container = False
            elif record_name.endswith('Blob'):
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

    def parse(self, more_data):
        instance, version = more_data
        if self.INSTANCE is not None and self.INSTANCE != instance:
            raise ValueError('invalid instance {0} for {1}'
                             .format(instance, self))
        elif self.INSTANCE is not None and instance not in (0,1):
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
        self.version =  version

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

    def parse(self, more_data):
        # set self.version and self.instance
        super(PptContainerRecord, self).parse(more_data)

        logging.debug('parsing contents of container record {0}'.format(self))

        # create a stream from self.data and parse it like any other
        data_stream = io.BytesIO(self.data)
        record_stream = PptStream(data_stream, self.size,
                                  'PptContainerRecordSubstream',
                                  record_base.STGTY_SUBSTREAM)
        self.records = list(record_stream.iter_records())
        logging.debug('done parsing contents of container record {0}'
                      .format(self))


class PptRecordCurrentUser(PptRecord):
    """ The CurrentUserAtom record """
    TYPE = 0x0ff6
    VERSION = 0
    INSTANCE = 0

    def parse(self, more_data):
        super(PptRecordCurrentUser, self).parse(more_data)
        if self.size < 24:
            raise ValueError('CurrentUser record is too small ({0})'
                             .format(self.size))
        self.size2, self.header_token, self.offset_to_current_edit, \
            self.len_user_name, self.doc_file_version, self.major_version, \
            self.minor_version, _ = unpack('<IIIHHBBH', self.data[0:20])
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
        self.release_version = unpack('<I', self.data[offset:offset+4])[0]
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
                            'stream'.format(len(data)))


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
    (0x1011, 'ExOleVbaActiveXAtom'),  # ExOleObj|VbaProject|ExControl]Stg[Unc|C]ompressedAtom
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
    # other types
    (0x0fc1, 'MetafileBlob'),
    (0x0fb8, 'FontEmbedDataBlob'),
    (0x07e7, 'SoundDataBlob'),
    (0x138b, 'BinaryTagDataBlob'),
])

# record types where version is not 0x0 or 0xf
VERSION_EXCEPTIONS = dict([
    (0x0400, 2),                       # rt_vbainfoatom
    (0x03ef, 2),                       # rt_slideatom
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


###############################################################################
# TESTING
###############################################################################


if __name__ == '__main__':
    def print_subrecords(record):
        if isinstance(record, PptContainerRecord):
            for subrec in record.records:
                logging.info('      {0}'.format(subrec))
        elif isinstance(record, PptRecordCurrentUser):
            logging.info('    crypt: {0}, offset {1}, user {2}/{3}'
                         .format(record.is_document_encrypted(),
                                 record.offset_to_current_edit,
                                 repr(record.ansi_user_name),
                                 repr(record.unicode_user_name)))
    sys.exit(record_base.test(sys.argv[1:], PptFile,
                              do_per_record=print_subrecords))
