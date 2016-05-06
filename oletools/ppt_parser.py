""" Parse a ppt (MS PowerPoint 97-2003) file

Based on olefile, parse the ppt-specific info

Code much influenced by olevba._extract_vba

Currently quite narrowly focused on extracting VBA from ppt files, no slides or
stuff, but built to be extended to parsing more/all of the file

References:
* https://msdn.microsoft.com/en-us/library/dd921564%28v=office.12%29.aspx
  and links there-in
"""

# === LICENSE =================================================================
# TODO
#------------------------------------------------------------------------------
# TODO:
# - make CurrentUserAtom and UserEditAtom PptTypes; adjust parse
# - make stream optional in PptUnexpectedData
# - license
# - create a AtomBase class that defines check_value and parses RecordHead?
#
# CHANGELOG:
# 2016-05-04 v0.01 CH: - start parsing "Current User" stream

__version__ = '0.01'


#--- IMPORTS ------------------------------------------------------------------
import sys
import logging
import struct
import traceback
import os

import thirdparty.olefile as olefile
from olevba import get_logger


# a global logger object used for debugging:
log = get_logger('ppt')


#--- CONSTANTS ----------------------------------------------------------------

# name of main stream
MAIN_STREAM_NAME = 'PowerPoint Document'

# URL and message to report issues:
URL_OLEVBA_ISSUES = 'https://bitbucket.org/decalage/oletools/issues'
MSG_OLEVBA_ISSUES = 'Please report this issue on %s' % URL_OLEVBA_ISSUES


# === EXCEPTIONS ==============================================================


class PptUnexpectedData(Exception):
    """ raise by PptParser if some field's value is not as expected """
    def __init__(self, stream, field_name, found_value, expected_value):
        self.msg = \
            'In stream "{}" for field "{}" found value "{}" but expected {}!' \
            .format(stream, field_name, found_value, expected_value)
        super(PptUnexpectedData, self).__init__(self.msg)


# === STRUCTS =================================================================


def check_value(name, value, expected):
    """ simplify verification of values in extract_from """
    if isinstance(expected, (list, tuple)):
        if value not in expected:
            exp_str = '[' + ' OR '.join('{0:04X}'.format(val)
                                        for val in expected) + ']'
            raise PptUnexpectedData(
                'Current User', name,
                '{0:04X}'.format(value), exp_str)
    elif expected != value:
        raise PptUnexpectedData(
            'Current User', name,
            '{0:04X}'.format(value), '{0:04X}'.format(expected))


class RecordHeader(object):
    """ a record header, at start of many types found in ppt files

    https://msdn.microsoft.com/en-us/library/dd926377%28v=office.12%29.aspx
    https://msdn.microsoft.com/en-us/library/dd948895%28v=office.12%29.aspx
    """

    def __init__(self):
        self.rec_ver = None
        self.rec_instance = None
        self.rec_type = None
        self.rec_len = None

    @classmethod
    def extract_from(clz, stream):
        """ reads 8 byte from stream """
        log.debug('parsing RecordHeader from stream')
        obj = clz()
        # first half byte is version, next 3 half bytes are instance
        version_instance, = struct.unpack('<H', stream.read(2))
        obj.rec_instance, obj.rec_ver = divmod(version_instance, 2**4)
        obj.rec_type, = struct.unpack('<H', stream.read(2))
        obj.rec_len, = struct.unpack('<L', stream.read(4))
        return obj


class CurrentUserAtom(object):
    """ An atom record that specifies information about the last user to modify
    the file and where the most recent user edit is located. This is the only
    record in the Current User Stream (section 2.1.1).

    https://msdn.microsoft.com/en-us/library/dd948895%28v=office.12%29.aspx
    """

    # allowed values for header_token
    HEADER_TOKEN_ENCRYPT = 0xF3D1C4DF
    HEADER_TOKEN_NOCRYPT = 0xE391C05F

    # allowed values for rel_version
    REL_VERSION_CAN_USE = 0x00000008
    REL_VERSION_NO_USE = 0x00000009

    # required values
    RECORD_TYPE = 0x0FF6
    SIZE = 0x14
    DOC_FILE_VERSION = 0x03F4
    MAJOR_VERSION = 0x03
    MINOR_VERSION = 0x00

    def __init__(self):
        self.rec_head = None
        self.size = None
        self.header_token = None
        self.offset_to_current_edit = None
        self.len_user_name = None
        self.doc_file_version = None
        self.major_version = None
        self.minor_version = None
        self.ansi_user_name = None
        self.unicode_user_name = None
        self.rel_version = None

    def is_encrypted(self):
        return self.header_token == self.HEADER_TOKEN_ENCRYPT

    @classmethod
    def extract_from(clz, ole):
        """ extract info from olefile """

        log.debug('parsing "Current User"')

        stream = None
        try:
            # open stream
            log.debug('opening stream')
            stream = ole.openstream('Current User')
            obj = clz()

            # parse record header
            obj.rec_head = RecordHeader.extract_from(stream)
            check_value('rec_version', obj.rec_head.rec_ver, 0)
            check_value('rec_instance', obj.rec_head.rec_instance, 0)
            check_value('rec_type', obj.rec_head.rec_type, clz.RECORD_TYPE)

            size, = struct.unpack('<L', stream.read(4))
            check_value('size', size, obj.SIZE)
            obj.header_token, = struct.unpack('<L', stream.read(4))
            check_value('headerToken', obj.header_token,
                        [clz.HEADER_TOKEN_ENCRYPT, clz.HEADER_TOKEN_NOCRYPT])
            obj.offset_to_current_edit, = struct.unpack('<L', stream.read(4))
            obj.len_user_name, = struct.unpack('<H', stream.read(2))
            if obj.len_user_name > 255:
                raise PptUnexpectedData(
                    'Current User', 'CurrentUserAtom.lenUserName',
                    obj.len_user_name, '< 256')
            obj.doc_file_version, = struct.unpack('<H', stream.read(2))
            check_value('docFileVersion', obj.doc_file_version,
                        clz.DOC_FILE_VERSION)
            obj.major_version, = struct.unpack('<B', stream.read(1))
            check_value('majorVersion', obj.major_version, clz.MAJOR_VERSION)
            obj.minor_version, = struct.unpack('<B', stream.read(1))
            check_value('minorVersion', obj.minor_version, clz.MINOR_VERSION)
            stream.read(2)    # unused
            obj.ansi_user_name = stream.read(obj.len_user_name)
            obj.rel_version, = struct.unpack('<L', stream.read(4))
            check_value('relVersion', obj.rel_version,
                        [clz.REL_VERSION_CAN_USE, clz.REL_VERSION_NO_USE])
            obj.unicode_user_name = stream.read(2 * obj.len_user_name)

            return obj

        finally:
            if stream is not None:
                log.debug('closing stream')
                stream.close()


class PptType(object):
    """ base class of data types found in ppt ole files

    starts with a RecordHeader, has a extract_from and a check_validity method
    """

    RECORD_TYPE = None      # must be specified in subclasses
    RECORD_INSTANCE = 0x0   # can be overwritten in subclasses
    RECORD_VERSION = 0x000  # can be overwritten in subclasses

    @classmethod
    def extract_from(clz, stream):
        raise NotImplementedError('abstract base function!')

    def __init__(self, stream_name=MAIN_STREAM_NAME):
        self.stream = None
        self.stream_name = stream_name
        self.rec_head = None

    def read_rec_head(self, stream):
        self.rec_head = RecordHeader.extract_from(stream)

    def set_stream(self, stream):
        """ need to call before any read_... method """
        self.stream = stream

    def unset_stream(self):
        """ should call after any read_... method """
        self.stream = None

    def read_1(self):
        """ read 1 byte from stream """
        return struct.unpack('<B', self.stream.read(1))[0]

    def read_2(self):
        """ read 2 byte (short) from stream """
        return struct.unpack('<H', self.stream.read(2))[0]

    def read_4(self):
        """ read 4 byte (long) from stream """
        return struct.unpack('<L', self.stream.read(4))[0]

    def check_validity(self):
        """ to be overwritten in subclasses

        :returns: list of PptUnexpectedData
        """
        raise NotImplementedError('abstract base function!')

    def check_value(self, name, value, expected):
        """ simplify verification of values: check value equals/is in expected

        :returns: list of PptUnexpectedData exceptions
        """
        if isinstance(expected, (list, tuple)):
            if value not in expected:
                clz_name = self.__class__.__name__
                exp_str = '[' + ' OR '.join('{0:04X}'.format(val)
                                            for val in expected) + ']'
                return [PptUnexpectedData(
                    self.stream_name, clz_name + '.' + name,
                    '{0:04X}'.format(value), exp_str), ]
        elif expected != value:
            clz_name = self.__class__.__name__
            return [PptUnexpectedData(
                self.stream_name, clz_name + '.' + name,
                '{0:04X}'.format(value), '{0:04X}'.format(expected)), ]
        return []

    def check_range(self, name, value, expect_lower, expect_upper):
        """ simplify verification of values: check value is in given range

        expect_lower or expected_upper can be given as None to check only one
        boundary. If value equals one of the boundaries, that is also an error
        (boundaries form an open interval)

        :returns: list of PptUnexpectedData exceptions
        """

        is_err = False
        if expect_upper is None and expect_lower is None:
            raise ValueError('need at least one non-None boundary!')
        if expect_lower is not None:
            if value <= expect_lower:
                is_err = True
        if expect_upper is not None:
            if value >= expect_upper:
                is_err = True

        if is_err:
            clz_name = self.__class__.__name__
            if expect_lower is None:
                expect_str = '< {0:04X}'.format(expect_upper)
            elif expect_upper is None:
                expect_str = '> {0:04X}'.format(expect_lower)
            else:
                expect_str = 'within ({0:04X}, {1:04X})'.format(expect_lower,
                                                                expect_upper)
            return [PptUnexpectedData(self.stream_name, clz_name + '.' + name,
                                      '{0:04X}'.format(value), expect_str), ]
        else:
            return []

    def check_rec_head(self, length=None):
        """ to be called by check_validity to check the self.rec_head

        uses self.RECORD_... constants, (not quite that constant for DummyType)
        """

        errs = []
        errs.extend(self.check_value('rec_head.recVer', self.rec_head.rec_ver,
                                     self.RECORD_VERSION))
        errs.extend(self.check_value('rec_head.recInstance',
                                     self.rec_head.rec_instance,
                                     self.RECORD_INSTANCE))
        if self.RECORD_TYPE is None:
            raise NotImplementedError('RECORD_TYPE not specified!')
        errs.extend(self.check_value('rec_head.recType',
                                     self.rec_head.rec_type,
                                     self.RECORD_TYPE))
        if length is not None:
            errs.extend(self.check_value('rec_head.recLen',
                                         self.rec_head.rec_len, length))
        return errs


class UserEditAtom(PptType):
    """ An atom record that specifies information about a user edit

    https://msdn.microsoft.com/en-us/library/dd945746%28v=office.12%29.aspx
    """

    RECORD_TYPE = 0x0FF5
    MINOR_VERSION = 0x00
    MAJOR_VERSION = 0x03

    def __init__(self):
        super(UserEditAtom, self).__init__()
        self.rec_head = None
        self.last_slide_id_ref = None
        self.version = None
        self.minor_version = None
        self.major_version = None
        self.offset_last_edit = None
        self.offset_persist_directory = None
        self.doc_persist_id_ref = None
        self.persist_id_seed = None
        self.last_view = None
        self.encrypt_session_persist_id_ref = None

    @classmethod
    def extract_from(clz, stream, is_encrypted):
        """ extract info from given stream (already positioned correctly!) """

        log.debug('extract UserEditAtom from stream')

        obj = clz()

        # parse record header
        obj.rec_head = RecordHeader.extract_from(stream)

        obj.last_slide_id_ref, = struct.unpack('<L', stream.read(4))
        obj.version, = struct.unpack('<H', stream.read(2))
        obj.minor_version, = struct.unpack('<B', stream.read(1))
        obj.major_version, = struct.unpack('<B', stream.read(1))
        obj.offset_last_edit, = struct.unpack('<L', stream.read(4))
        obj.offset_persist_directory, = struct.unpack('<L', stream.read(4))
        obj.doc_persist_id_ref, = struct.unpack('<L', stream.read(4))
        obj.persist_id_seed, = struct.unpack('<L', stream.read(4))
        # (can only check once have the PersistDirectoryAtom)
        obj.last_view, = struct.unpack('<H', stream.read(2))
        stream.read(2)   # unused
        if is_encrypted:
            obj.encrypt_session_persist_id_ref, = \
                struct.unpack('<L', stream.read(4))
        else:   # this entry may be there or may not
            obj.encrypt_session_persist_id_ref = None

        return obj

    def check_validity(self, offset=None):
        errs = self.check_rec_head()
        errs.extend(self.check_value('minorVersion', self.minor_version,
                                     self.MINOR_VERSION))
        errs.extend(self.check_value('majorVersion', self.major_version,
                                     self.MAJOR_VERSION))
        if offset is not None:
            if self.offset_last_edit >= offset:
                errs.append(PptUnexpectedData(
                    'PowerPoint Document', 'UserEditAtom.offsetLastEdit',
                    self.offset_last_edit, '< {}'.format(offset)))
            if self.offset_persist_directory >= offset or \
                    self.offset_persist_directory <= self.offset_last_edit:
                errs.append(PptUnexpectedData(
                    'PowerPoint Document',
                    'UserEditAtom.offsetPersistDirectory',
                    self.offset_last_edit,
                    'in ({}, {})'.format(self.offset_last_edit, offset)))
        errs.extend(self.check_value('docPersistIdRef',
                                     self.doc_persist_id_ref, 1))
        return errs

        # TODO: offer to check persist_id_seed given PersistDirectoryAtom)


class DummyType(PptType):
    """ a type that is found in ppt documents we are not interested in

    instead of parsing many uninteresting types, we just read their
    RecordHeader and set the RECORD_... values on an instance- (instead of
    class-) level

    used to skip over uninteresting types in e.g. DocumentContainer
    """

    def __init__(self, type_name, record_type, rec_ver=0, rec_instance=0,
                 rec_len=None):
        super(DummyType, self).__init__()
        self.type_name = type_name
        self.RECORD_TYPE = record_type
        self.RECORD_VERSION = rec_ver
        self.RECORD_INSTANCE = rec_instance
        self.record_length = rec_len

    def extract_from(self, stream):
        """ extract record header and just skip as many bytes as header says

        Since this requires RECORD_... values set in constructor, this is NOT
        a classmethod like all the other extract_from!

        Otherwise this tries to be compatible with other extract_from methods
        (e.g. returns self)
        """
        self.read_rec_head(stream)
        log.debug('skipping over {} Byte for type {}'
                  .format(self.rec_head.rec_len, self.type_name))
        log.debug('start at pos {}'.format(stream.tell()))
        stream.seek(self.rec_head.rec_len, os.SEEK_CUR)
        log.debug('now at pos {}'.format(stream.tell()))
        return self

    def check_validity(self):
        return self.check_rec_head(self.record_length)


class PersistDirectoryAtom(PptType):
    """ one part of a persist object directory with unique persist object id

    contains PersistDirectoryEntry objects

    https://msdn.microsoft.com/en-us/library/dd952680%28v=office.12%29.aspx
    """

    RECORD_TYPE = 0x1772

    def __init__(self):
        super(PersistDirectoryAtom, self).__init__()
        self.rg_persist_dir_entry = None    # actually, this will be an array
        self.stream_offset = None

    @classmethod
    def extract_from(clz, stream):
        """ create and return object with data from given stream """

        log.debug("Extracting a PersistDirectoryAtom from stream")
        obj = clz()

        # remember own offset for checking validity
        obj.stream_offset = stream.tell()

        # parse record header
        obj.read_rec_head(stream)

        # read directory entries from list until reach size for this object
        curr_pos = stream.tell()
        stop_pos = curr_pos + obj.rec_head.rec_len
        log.debug('start reading at pos {}, read until {}'
                  .format(curr_pos, stop_pos))
        obj.rg_persist_dir_entry = []

        while curr_pos < stop_pos:
            new_entry = PersistDirectoryEntry.extract_from(stream)
            obj.rg_persist_dir_entry.append(new_entry)
            curr_pos = stream.tell()
            log.debug('at pos {}'.format(curr_pos))
        return obj

    def check_validity(self, user_edit_last_offset=None):
        errs = self.check_rec_head()
        for entry in self.rg_persist_dir_entry:
            errs.extend(entry.check_validity(user_edit_last_offset,
                                             self.stream_offset))
        return errs


class PersistDirectoryEntry(object):
    """ an entry contained in a PersistDirectoryAtom.rg_persist_dir_entry

    A structure that specifies a compressed table of sequential persist object
    identifiers and stream offsets to associated persist objects.

    NOT a subclass of PptType because has no RecordHeader

    https://msdn.microsoft.com/en-us/library/dd947347%28v=office.12%29.aspx
    """

    def __init__(self):
        self.persist_id = None
        self.c_persist = None
        self.rg_persist_offset = None

    @classmethod
    def extract_from(clz, stream):
        # take a 4-byte (=32bit) number, divide into 20bit and 12 bit)
        log.debug("Extracting a PersistDirectoryEntry from stream")
        obj = clz()

        # persistId (20 bits): An unsigned integer that specifies a starting
        # persist object identifier. It MUST be less than or equal to 0xFFFFE.
        # The first entry in rgPersistOffset is associated with persistId. The
        # next entry, if present, is associated with persistId plus 1. Each
        # entry in rgPersistOffset is associated with a persist object
        # identifier in this manner, with the final entry associated with
        # persistId + cPersist - 1.

        # cPersist (12 bits): An unsigned integer that specifies the count of
        # items in rgPersistOffset. It MUST be greater than or equal to 0x001.
        temp, = struct.unpack('<L', stream.read(4))
        obj.c_persist, obj.persist_id = divmod(temp, 2**20)
        log.debug('temp is 0x{0:04X} --> id is {1}, reading {2} offsets'
                  .format(temp, obj.persist_id, obj.c_persist))

        # rgPersistOffset (variable): An array of PersistOffsetEntry (section
        # 2.3.6) that specifies stream offsets to persist objects. The count of
        # items in the array is specified by cPersist. The value of each item
        # MUST be greater than or equal to offsetLastEdit in the corresponding
        # user edit and MUST be less than the offset, in bytes, of the
        # corresponding persist object directory.
        # PersistOffsetEntry: An unsigned 4-byte integer that specifies an
        # offset, in bytes, from the beginning of the PowerPoint Document
        # Stream (section 2.1.2) to a persist object.
        obj.rg_persist_offset = [struct.unpack('<L', stream.read(4))[0] \
                                 for _ in range(obj.c_persist)]
        log.debug('offsets are: {}'.format(obj.rg_persist_offset))
        return obj

    def check_validity(self, user_edit_last_offset=None,
                       persist_obj_dir_offset=None):
        errs = []
        if self.persist_id > 0xFFFFE:  # (--> == 0xFFFFF since 20bit)
            errs.append(PptUnexpectedData(
                MAIN_STREAM_NAME, 'PersistDirectoryEntry.persist_id',
                self.persist_id, '< 0xFFFFE (dec: {})'.format(0xFFFFE)))
        if self.c_persist == 0:
            errs.append(PptUnexpectedData(
                MAIN_STREAM_NAME, 'PersistDirectoryEntry.c_persist',
                self.c_persist, '> 0'))
        if user_edit_last_offset is not None \
                and min(self.rg_persist_offset) < user_edit_last_offset:
            errs.append(PptUnexpectedData(
                MAIN_STREAM_NAME, 'PersistDirectoryEntry.rg_persist_offset',
                min(self.rg_persist_offset),
                '> UserEdit.offsetLastEdit ({})'
                .format(user_edit_last_offset)))
        if persist_obj_dir_offset is not None \
                and max(self.rg_persist_offset) > persist_obj_dir_offset:
            errs.append(PptUnexpectedData(
                MAIN_STREAM_NAME, 'PersistDirectoryEntry.rg_persist_offset',
                max(self.rg_persist_offset),
                '> PersistObjectDirectory offset ({})'
                .format(persist_obj_dir_offset)))
        return errs


class DocInfoListContainer(PptType):
    """ information about the document and document display settings

    https://msdn.microsoft.com/en-us/library/dd926767%28v=office.12%29.aspx
    """

    RECORD_VERSION = 0xF
    RECORD_TYPE = 0x07D0

    def __init__(self):
        super(DocInfoListContainer, self).__init__()


class DocumentContainer(PptType):
    """ a DocumentContainer record

    https://msdn.microsoft.com/en-us/library/dd947357%28v=office.12%29.aspx
    """

    RECORD_TYPE = 0x03E8

    def __init__(self):
        super(DocumentContainer, self).__init__()
        self.document_atom = None
        self.ex_obj_list = None
        self.document_text_info = None
        self.sound_collection = None
        self.drawing_group = None
        self.master_list = None
        self.doc_info_list = None
        self.slide_hf = None
        self.notes_hf = None
        self.slide_list = None
        self.notes_list = None
        self.slide_show_doc_info = None
        self.named_shows = None
        self.summary = None
        self.doc_routing_slip = None
        self.print_options = None
        self.rt_custom_table_styles_1 = None
        self.end_document = None
        self.rt_custom_table_styles_2 = None

    @classmethod
    def extract_from(clz, stream):
        """ created object with values from given stream

        stream is assumed to be positioned correctly

        this container contains lots of data we are not interested in.
        """
        obj = clz()

        # parse record header
        obj.read_rec_head(stream)

        # documentAtom (48 bytes): A DocumentAtom record (section 2.4.2) that
        # specifies size information for presentation slides and notes slides.
        obj.document_atom = DummyType('DocumentAtom', 0x03E9, rec_ver=0x1,
                                      rec_len=0x28).extract_from(stream)

        # exObjList (variable): An optional ExObjListContainer record (section
        # 2.10.1) that specifies the list of external objects in the document.
        obj.ex_obj_list = DummyType('ExObjListContainer', 0x0409, rec_ver=0xF)\
                          .extract_from(stream)

        # documentTextInfo (variable): A DocumentTextInfoContainer record
        # (section 2.9.1) that specifies the default text styles for the
        # document.
        obj.document_text_info = DummyType('DocumentTextInfoContainer', 0x03F2,
                                           rec_ver=0xF).extract_from(stream)

        # soundCollection (variable): An optional SoundCollectionContainer
        # record (section 2.4.16.1) that specifies the list of sounds in the
        # file.
        obj.sound_collection = DummyType('SoundCollectionContainer', 0x07E4,
                                         rec_ver=0xF, rec_instance=0x005)\
                               .extract_from(stream)

        # drawingGroup (variable): A DrawingGroupContainer record (section
        # 2.4.3) that specifies drawing information for the document.
        obj.drawing_group = DummyType('DrawingGroupContainer', 0x040B,
                                      rec_ver=0xF).extract_from(stream)

        # masterList (variable): A MasterListWithTextContainer record (section
        # 2.4.14.1) that specifies the list of main master slides and title
        # master slides.
        obj.master_list = DummyType('MasterListWithContainer', 0x0FF0,
                                    rec_ver=0xF).extract_from(stream)

        # docInfoList (variable): An optional DocInfoListContainer record
        # (section 2.4.4) that specifies additional document information.
        # this is the variable we are interested in!
        obj.doc_info_list = DocInfoListContainer.extract_from(stream)

        # slideHF (variable): An optional SlideHeadersFootersContainer record
        # (section 2.4.15.1) that specifies the default header and footer
        # information for presentation slides.
        obj.slide_hf = None

        # notesHF (variable): An optional NotesHeadersFootersContainer record
        # (section 2.4.15.6) that specifies the default header and footer
        # information for notes slides.
        obj.notes_hf = None

        # slideList (variable): An optional SlideListWithTextContainer record
        # (section 2.4.14.3) that specifies the list of presentation slides.
        obj.slide_list = None

        # notesList (variable): An optional NotesListWithTextContainer record
        # (section 2.4.14.6) that specifies the list of notes slides.
        obj.notes_list = None

        # slideShowDocInfoAtom (88 bytes): An optional SlideShowDocInfoAtom
        # record (section 2.6.1) that specifies slide show information for the
        # document.
        obj.slide_show_doc_info = None

        # namedShows (variable): An optional NamedShowsContainer record
        # (section 2.6.2) that specifies named shows in the document.
        obj.named_shows = None

        # summary (variable): An optional SummaryContainer record (section
        # 2.4.22.3) that specifies bookmarks for the document.
        obj.summary = None

        # docRoutingSlipAtom (variable): An optional DocRoutingSlipAtom record
        # (section 2.11.1) that specifies document routing information.
        obj.doc_routing_slip = None

        # printOptionsAtom (13 bytes): An optional PrintOptionsAtom record
        # (section 2.4.12) that specifies default print options.
        obj.print_options = None

        # rtCustomTableStylesAtom1 (variable): An optional
        # RoundTripCustomTableStyles12Atom record (section 2.11.13) that
        # specifies round-trip information for custom table styles.
        obj.rt_custom_table_styles_1 = None

        # endDocumentAtom (8 bytes): An EndDocumentAtom record (section 2.4.13)
        # that specifies the end of the information for the document.
        obj.end_document = None

        # rtCustomTableStylesAtom2 (variable): An optional
        # RoundTripCustomTableStyles12Atom record that specifies round-trip
        # information for custom table styles. It MUST NOT exist if
        # rtCustomTableStylesAtom1 exists.
        obj.rt_custom_table_styles_2 = None

        return obj


    def check_validity(self):
        """ check all values in object for valid values """
        errs = self.check_rec_head()
        errs.extend(self.document_atom.check_validity())
        errs.extend(self.ex_obj_list.check_validity())
        errs.extend(self.document_text_info.check_validity())
        errs.extend(self.sound_collection.check_validity())
        errs.extend(self.drawing_group.check_validity())
        errs.extend(self.master_list.check_validity())
        errs.extend(self.doc_info_list.check_validity())
        return errs

# === PptParser ===============================================================


class PptParser(object):
    """ Parser for PowerPoint 97-2003 specific data structures

    requires an OleFileIO
    """

    def __init__(self, ole, fast_fail=False):
        """ constructor

        :param ole: OleFileIO or anything that OleFileIO constructor accepts
        :param bool fast_fail: if True, all unexpected data will raise a
                               PptUnexpectedData; if False will only log error
        """
        if isinstance(ole, olefile.OleFileIO):
            self.ole = ole
        else:
            log.debug('Opening file ' + ole)
            self.ole = olefile.OleFileIO(ole)

        self.fast_fail = fast_fail

        self.current_user_atom = None
        self.document_persist_obj = None
        self.persist_object_directory = None

        # basic compatibility check: root directory structure is
        # [['\x05DocumentSummaryInformation'],
        #  ['\x05SummaryInformation'],
        #  ['Current User'],
        #  ['PowerPoint Document']]
        root_streams = self.ole.listdir()
        for stream in root_streams:
            log.debug('found root stream {!r}'.format(stream))
        if any(len(stream) != 1 for stream in root_streams):
            self._fail('root', 'listdir', root_streams, 'len = 1')
        root_streams = [stream[0].lower() for stream in root_streams]
        if not 'current user' in root_streams:
            self._fail('root', 'listdir', root_streams, 'Current User')
        if not MAIN_STREAM_NAME.lower() in root_streams:
            self._fail('root', 'listdir', root_streams, MAIN_STREAM_NAME)

    def _log_exception(self, msg=None):
        """ log an exception instead of raising it

        call in one of 2 ways:
            try:
                if fail():
                    self._log_exception('this is the message')
            except:
                self._log_exception()   # only possible in except clause
        """
        if msg is not None:
            stack = traceback.extract_stack()[:-1]
        else:
            _, exc, trace = sys.exc_info()
            stack = traceback.extract_tb(trace)
            msg = str(exc)
        log.error(msg)

        for i_entry, entry in enumerate(traceback.format_list(stack)):
            for line in entry.splitlines():
                log.debug('trace {}: {}'.format(i_entry, line))

    def _fail(self, *args):
        """ depending on self.fast_fail raise PptUnexpectedData or just log err

        args as for PptUnexpectedData
        """
        if self.fast_fail:
            raise PptUnexpectedData(*args)
        else:
            self._log_exception(PptUnexpectedData(*args).msg)

    def parse_current_user(self):
        """ parse the CurrentUserAtom record from stream 'Current User'

        Structure described in
        https://msdn.microsoft.com/en-us/library/dd948895%28v=office.12%29.aspx
        """

        if self.current_user_atom is not None:
            log.warning('re-reading and overwriting '
                        'previously read current_user_atom')

        try:
            self.current_user_atom = CurrentUserAtom.extract_from(self.ole)
        except Exception:
            if self.fast_fail:
                raise
            else:
                self._log_exception()

    def parse_persist_object_directory(self):
        """ Part 1: Construct the persist object directory """

        if self.persist_object_directory is not None:
            log.warning('re-reading and overwriting '
                        'previously read persist_object_directory')

        if self.current_user_atom is None:
            self.parse_current_user()

        offset = self.current_user_atom.offset_to_current_edit
        is_encrypted = self.current_user_atom.is_encrypted()
        self.persist_object_directory = {}

        stream = None
        try:
            log.debug('opening stream')
            stream = self.ole.openstream(MAIN_STREAM_NAME)
            while offset != 0:

                stream.seek(offset, os.SEEK_SET)
                user_edit = UserEditAtom.extract_from(stream, is_encrypted)

                log.debug('checking validity')
                errs = user_edit.check_validity()
                if errs:
                    log.warning('check_validity found {} issues'
                                .format(len(errs)))
                for err in errs:
                    log.warning('UserEditAtom.check_validity: {}'.format(err))
                if errs and self.fast_fail:
                    raise errs[0]

                log.debug('seeking to pos {}'
                          .format(user_edit.offset_persist_directory))
                stream.seek(user_edit.offset_persist_directory, os.SEEK_SET)

                persist_dir_atom = PersistDirectoryAtom.extract_from(stream)

                log.debug('checking validity')
                errs = persist_dir_atom.check_validity(offset)
                if errs:
                    log.warning('check_validity found {} issues'
                                .format(len(errs)))
                for err in errs:
                    log.warning('PersistDirectoryAtom.check_validity: {}'
                                .format(err))
                if errs and self.fast_fail:
                    raise errs[0]

                for entry in persist_dir_atom.rg_persist_dir_entry:
                    log.debug('saving {} offsets for persist_id {}'
                              .format(len(entry.rg_persist_offset),
                                      entry.persist_id))
                    self.persist_object_directory[entry.persist_id] = \
                        entry.rg_persist_offset

                # check for more
                offset = user_edit.offset_last_edit
        except Exception:
            if self.fast_fail:
                raise
            else:
                self._log_exception()
        finally:
            if stream is not None:
                log.debug('closing stream')
                stream.close()

    def parse_document_persist_object(self):
        """ """
        if self.document_persist_obj is not None:
            log.warning('re-reading and overwriting '
                        'previously read document_persist_object')

        if self.persist_object_directory is None:
            self.parse_persist_object_directory()

        offset = None  # TODO: read from object directory
        stream = None

        try:
            log.debug('opening stream')
            stream = self.ole.openstream(MAIN_STREAM_NAME)
            log.debug('stream pos: {}'.format(stream.tell()))
            stream.seek(offset)
            log.debug('seek by {} to {}'.format(offset, stream.tell()))
            self.document_persist_obj = DocumentContainer.extract_from(stream)
        except Exception:
            if self.fast_fail:
                raise
            else:
                self._log_exception()
        finally:
            if stream is not None:
                log.debug('closing stream')
                stream.close()

        log.debug('checking validity')
        errs = self.document_persist_obj.check_validity()
        if errs:
            log.warning('check_validity found {} issues'.format(len(errs)))
        for err in errs:
            log.warning('check_validity(document_persist_obj): {}'
                        .format(err))
        if errs and self.fast_fail:
            raise errs[0]

# === TESTING =================================================================

def test():
    """ for testing and debugging """

    # setup logging
    logging.basicConfig(level=logging.DEBUG,
                        format='%(levelname)-8s %(message)s')
    log.setLevel(logging.NOTSET)

    # test file with some autostart macros
    test_file = 'gelaber_autostart.ppt'

    # parse
    ppt = PptParser(test_file, fast_fail=False)
    ppt.parse_document_persist_object()


if __name__ == '__main__':
    test()
