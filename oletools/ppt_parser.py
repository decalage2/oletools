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
# - license
#
# CHANGELOG:
# 2016-05-04 v0.01 CH: - start parsing "Current User" stream

__version__ = '0.01'


#--- IMPORTS ------------------------------------------------------------------
import sys
import logging
import struct
import traceback

import thirdparty.olefile as olefile
from olevba import get_logger


# a global logger object used for debugging:
log = get_logger('ppt')


#--- CONSTANTS ----------------------------------------------------------------


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


class RecordHeader(object):
    """ a record header, often found in ppt files

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
        obj.rec_ver, obj.rec_instance = divmod(version_instance, 16)
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

    @classmethod
    def extract_from(clz, ole):
        """ extract info from olefile """

        log.debug('parsing "Current User"')

        stream = None
        try:
            # open stream
            stream = ole.openstream('Current User')
            obj = clz()

            # parse record header
            obj.rec_head = RecordHeader.extract_from(stream)
            obj.check_value('rec_version', obj.rec_head.rec_ver, 0)
            obj.check_value('rec_instance', obj.rec_head.rec_ver, 0)
            obj.check_value('rec_instance', obj.rec_head.rec_type,
                             clz.RECORD_TYPE)

            size, = struct.unpack('<L', stream.read(4))
            obj.check_value('size', size, obj.SIZE)
            obj.header_token, = struct.unpack('<L', stream.read(4))
            obj.check_value('headerToken', obj.header_token,
                             [clz.HEADER_TOKEN_ENCRYPT,
                              clz.HEADER_TOKEN_NOCRYPT])
            log.debug('headerToken is encrypt: {}'
                      .format(obj.header_token == clz.HEADER_TOKEN_ENCRYPT))
            obj.offset_to_current_edit, = struct.unpack('<L', stream.read(4))
            log.debug('offsetToCurrentEdit: {0} ({0:04X})'
                      .format(obj.offset_to_current_edit))
            obj.len_user_name, = struct.unpack('<H', stream.read(2))
            log.debug('lenUserName: {}'.format(obj.len_user_name))
            if obj.len_user_name > 255:
                raise PptUnexpectedData(
                    'Current User', 'CurrentUserAtom.lenUserName',
                    obj.len_user_name, '< 256')
            obj.doc_file_version, = struct.unpack('<H', stream.read(2))
            obj.check_value('docFileVersion', obj.doc_file_version,
                             clz.DOC_FILE_VERSION)
            obj.major_version, = struct.unpack('<B', stream.read(1))
            obj.check_value('majorVersion', obj.major_version,
                             clz.MAJOR_VERSION)
            obj.minor_version, = struct.unpack('<B', stream.read(1))
            obj.check_value('minorVersion', obj.minor_version,
                             clz.MINOR_VERSION)
            stream.read(2)    # unused
            obj.ansi_user_name = stream.read(obj.len_user_name)
            log.debug('ansiUserName: {!r}'.format(obj.ansi_user_name))
            obj.rel_version, = struct.unpack('<L', stream.read(4))
            log.debug('relVersion: {0:04X}'.format(obj.rel_version))
            obj.check_value('relVersion', obj.rel_version,
                             [clz.REL_VERSION_CAN_USE,
                              clz.REL_VERSION_NO_USE])
            obj.unicode_user_name = stream.read(2 * obj.len_user_name)
            log.debug('unicodeUserName: {!r}'.format(obj.unicode_user_name))

            return obj

        except Exception:
            raise
        finally:
            if stream is not None:
                log.debug('closing stream')
                stream.close()

    def check_value(self, name, value, expected):
        """ simplify verification of values in extract_from """
        if isinstance(expected, (list, tuple)):
            if value not in expected:
                exp_str = '[' + ' OR '.join('{0:04X}'.format(val) 
                                            for val in expected) + ']'
                raise PptUnexpectedData(
                    'Current User', 'CurrentUserAtom.' + name,
                    '{0:04X}'.format(value), exp_str)
        elif expected != value:
            raise PptUnexpectedData(
                'Current User', 'CurrentUserAtom.' + name,
                '{0:04X}'.format(value), '{0:04X}'.format(expected))


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
        if not 'powerpoint document' in root_streams:
            self._fail('root', 'listdir', root_streams, 'PowerPoint Document')

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

        try:
            self.current_user_atom = CurrentUserAtom.extract_from(self.ole)
        except Exception:
            if self.fast_fail:
                raise
            else:
                self._log_exception()

# === TESTING =================================================================

def test():
    """ for testing and debugging """

    # setup logging
    logging.basicConfig(level=logging.DEBUG, format='%(levelname)-8s %(message)s')
    log.setLevel(logging.NOTSET)

    # test file with some autostart macros
    test_file = 'gelaber_autostart.ppt'

    # parse
    ppt = PptParser(test_file)
    ppt.parse_current_user()


if __name__ == '__main__':
    test()
