#!/usr/bin/env python
"""
oleobj.py

oleobj is a Python script and module to parse OLE objects and files stored
into various MS Office file formats (doc, xls, ppt, docx, xlsx, pptx, etc)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

oleobj is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE =================================================================

# oleobj is copyright (c) 2015-2022 Philippe Lagadec (http://www.decalage.info)
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


# -- IMPORTS ------------------------------------------------------------------

from __future__ import print_function

import logging
import struct
import argparse
import os
import re
import sys
import io
from zipfile import is_zipfile
import random

import olefile

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty import xglob
from oletools.ppt_record_parser import (is_ppt, PptFile,
                                        PptRecordExOleVbaActiveXAtom)
from oletools.ooxml import XmlParser
from oletools.common.io_encoding import ensure_stdout_handles_unicode

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2015-12-05 v0.01 PL: - first version
# 2016-06          PL: - added main and process_file (not working yet)
# 2016-07-18 v0.48 SL: - added Python 3.5 support
# 2016-07-19       PL: - fixed Python 2.6-7 support
# 2016-11-17 v0.51 PL: - fixed OLE native object extraction
# 2016-11-18       PL: - added main for setup.py entry point
# 2017-05-03       PL: - fixed absolute imports (issue #141)
# 2018-01-18 v0.52 CH: - added support for zipped-xml-based types (docx, pptx,
#                        xlsx), and ppt
# 2018-03-27       PL: - fixed issue #274 in read_length_prefixed_string
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2018-10-30       SA: - added detection of external links (PR #317)
# 2020-03-03 v0.56 PL: - fixed bug #541, "Ole10Native" is case-insensitive
# 2022-01-28 v0.60 PL: - added detection of customUI tags

__version__ = '0.60.1'

# -----------------------------------------------------------------------------
# TODO:
# + setup logging (common with other oletools)


# -----------------------------------------------------------------------------
# REFERENCES:

# Reference for the storage of embedded OLE objects/files:
# [MS-OLEDS]: Object Linking and Embedding (OLE) Data Structures
# https://msdn.microsoft.com/en-us/library/dd942265.aspx

# - officeparser: https://github.com/unixfreak0037/officeparser
# TODO: oledump


# === LOGGING =================================================================

DEFAULT_LOG_LEVEL = "warning"
LOG_LEVELS = {'debug':    logging.DEBUG,
              'info':     logging.INFO,
              'warning':  logging.WARNING,
              'error':    logging.ERROR,
              'critical': logging.CRITICAL,
              'debug-olefile': logging.DEBUG}


class NullHandler(logging.Handler):
    """
    Log Handler without output, to avoid printing messages if logging is not
    configured by the main application.
    Python 2.7 has logging.NullHandler, but this is necessary for 2.6:
    see https://docs.python.org/2.6/library/logging.html section
    configuring-logging-for-a-library
    """
    def emit(self, record):
        pass


def get_logger(name, level=logging.CRITICAL+1):
    """
    Create a suitable logger object for this module.
    The goal is not to change settings of the root logger, to avoid getting
    other modules' logs on the screen.
    If a logger exists with same name, reuse it. (Else it would have duplicate
    handlers and messages would be doubled.)
    The level is set to CRITICAL+1 by default, to avoid any logging.
    """
    # First, test if there is already a logger with the same name, else it
    # will generate duplicate messages (due to duplicate handlers):
    if name in logging.Logger.manager.loggerDict:
        # NOTE: another less intrusive but more "hackish" solution would be to
        # use getLogger then test if its effective level is not default.
        logger = logging.getLogger(name)
        # make sure level is OK:
        logger.setLevel(level)
        return logger
    # get a new logger:
    logger = logging.getLogger(name)
    # only add a NullHandler for this logger, it is up to the application
    # to configure its own logging:
    logger.addHandler(NullHandler())
    logger.setLevel(level)
    return logger


# a global logger object used for debugging:
log = get_logger('oleobj')     # pylint: disable=invalid-name


def enable_logging():
    """
    Enable logging for this module (disabled by default).
    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    log.setLevel(logging.NOTSET)


# === CONSTANTS ===============================================================

# some str methods on Python 2.x return characters,
# while the equivalent bytes methods return integers on Python 3.x:
if sys.version_info[0] <= 2:
    # Python 2.x
    NULL_CHAR = '\x00'
else:
    # Python 3.x
    NULL_CHAR = 0
    xrange = range    # pylint: disable=redefined-builtin, invalid-name

OOXML_RELATIONSHIP_TAG = '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'
# There are several customUI tags for different versions of Office:
TAG_CUSTOMUI_2007 = "{http://schemas.microsoft.com/office/2006/01/customui}customUI"
TAG_CUSTOMUI_2010 = "{http://schemas.microsoft.com/office/2009/07/customui}customUI"

# === GLOBAL VARIABLES ========================================================

# struct to parse an unsigned integer of 32 bits:
STRUCT_UINT32 = struct.Struct('<L')
assert STRUCT_UINT32.size == 4  # make sure it matches 4 bytes

# struct to parse an unsigned integer of 16 bits:
STRUCT_UINT16 = struct.Struct('<H')
assert STRUCT_UINT16.size == 2  # make sure it matches 2 bytes

# max length of a zero-terminated ansi string. Not sure what this really is
STR_MAX_LEN = 1024

# size of chunks to copy from ole stream to file
DUMP_CHUNK_SIZE = 4096

# return values from main; can be added
# (e.g.: did dump but had err parsing and dumping --> return 1+4+8 = 13)
RETURN_NO_DUMP = 0     # nothing found to dump/extract
RETURN_DID_DUMP = 1    # did dump/extract successfully
RETURN_ERR_ARGS = 2    # reserve for OptionParser.parse_args
RETURN_ERR_STREAM = 4  # error opening/parsing a stream
RETURN_ERR_DUMP = 8    # error dumping data from stream to file

# Not sure if they can all be "External", but just in case
BLACKLISTED_RELATIONSHIP_TYPES = [
    'attachedTemplate',
    'externalLink',
    'externalLinkPath',
    'externalReference',
    'frame',
    'hyperlink',
    'officeDocument',
    'oleObject',
    'package',
    'slideUpdateUrl',
    'slideMaster',
    'slide',
    'slideUpdateInfo',
    'subDocument',
    'worksheet'
]

# Save maximum length of a filename
MAX_FILENAME_LENGTH = 255

# Max attempts at generating a non-existent random file name
MAX_FILENAME_ATTEMPTS = 100

# === FUNCTIONS ===============================================================


def read_uint32(data, index):
    """
    Read an unsigned integer from the first 32 bits of data.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index to start reading from or None if data is stream.
    :return: tuple (value, index) containing the read value (int),
             and the index to continue reading next time.
    """
    if index is None:
        value = STRUCT_UINT32.unpack(data.read(4))[0]
    else:
        value = STRUCT_UINT32.unpack(data[index:index+4])[0]
        index += 4
    return (value, index)


def read_uint16(data, index):
    """
    Read an unsigned integer from the 16 bits of data following index.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index to start reading from or None if data is stream
    :return: tuple (value, index) containing the read value (int),
             and the index to continue reading next time.
    """
    if index is None:
        value = STRUCT_UINT16.unpack(data.read(2))[0]
    else:
        value = STRUCT_UINT16.unpack(data[index:index+2])[0]
        index += 2
    return (value, index)


def read_length_prefixed_string(data, index):
    """
    Read a length-prefixed ANSI string from data.

    :param data: bytes string or stream containing the data to be extracted.
    :param index: index in data where string size start or None if data is
                  stream
    :return: tuple (value, index) containing the read value (bytes string),
             and the index to start reading from next time.
    """
    length, index = read_uint32(data, index)
    # if length = 0, return a null string (no null character)
    if length == 0:
        return ('', index)
    # extract the string without the last null character
    if index is None:
        ansi_string = data.read(length-1)
        null_char = data.read(1)
    else:
        ansi_string = data[index:index+length-1]
        null_char = data[index+length-1]
        index += length
    # TODO: only in strict mode:
    # check the presence of the null char:
    assert null_char == NULL_CHAR
    return (ansi_string, index)


def guess_encoding(data):
    """ guess encoding of byte string to create unicode

    Since this is used to decode path names from ole objects, prefer latin1
    over utf* codecs if ascii is not enough
    """
    for encoding in 'ascii', 'latin1', 'utf8', 'utf-16-le', 'utf16':
        try:
            result = data.decode(encoding, errors='strict')
            log.debug(u'decoded using {0}: "{1}"'.format(encoding, result))
            return result
        except UnicodeError:
            pass
    log.warning('failed to guess encoding for string, falling back to '
                'ascii with replace')
    return data.decode('ascii', errors='replace')


def read_zero_terminated_string(data, index):
    """
    Read a zero-terminated string from data

    :param data: bytes string or stream containing an ansi string
    :param index: index at which the string should start or None if data is
                  stream
    :return: tuple (unicode, index) containing the read string (unicode),
             and the index to start reading from next time.
    """
    if index is None:
        result = bytearray()
        # pylint: disable-next=possibly-used-before-assignment
        for _ in xrange(STR_MAX_LEN):
            char = ord(data.read(1))    # need ord() for py3
            if char == 0:
                return guess_encoding(result), index
            result.append(char)
        raise ValueError('found no string-terminating zero-byte!')
    else:       # data is byte array, can just search
        end_idx = data.index(b'\x00', index, index+STR_MAX_LEN)
        # encode and return with index after the 0-byte
        return guess_encoding(data[index:end_idx]), end_idx+1


# === CLASSES =================================================================


class OleNativeStream(object):
    """
    OLE object contained into an OLENativeStream structure.
    (see MS-OLEDS 2.3.6 OLENativeStream)

    Filename and paths are decoded to unicode.
    """
    # constants for the type attribute:
    # see MS-OLEDS 2.2.4 ObjectHeader
    TYPE_LINKED = 0x01
    TYPE_EMBEDDED = 0x02

    def __init__(self, bindata=None, package=False):
        """
        Constructor for OleNativeStream.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: forwarded to parse, see docu there
        :param package: bool, set to True when extracting from an OLE Package
                        object
        """
        self.filename = None
        self.src_path = None
        self.unknown_short = None
        self.unknown_long_1 = None
        self.unknown_long_2 = None
        self.temp_path = None
        self.actual_size = None
        self.data = None
        self.package = package
        self.is_link = None
        self.data_is_stream = None
        if bindata is not None:
            self.parse(data=bindata)

    def parse(self, data):
        """
        Parse binary data containing an OLENativeStream structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.3.6 OLENativeStream)

        :param data: bytes array or stream, containing OLENativeStream
                     structure containing an OLE object
        :return: None
        """
        # TODO: strict mode to raise exceptions when values are incorrect
        # (permissive mode by default)
        if hasattr(data, 'read'):
            self.data_is_stream = True
            index = None       # marker for read_* functions to expect stream
        else:
            self.data_is_stream = False
            index = 0          # marker for read_* functions to expect array

        # An OLE Package object does not have the native data size field
        if not self.package:
            self.native_data_size, index = read_uint32(data, index)
            log.debug('OLE native data size = {0:08X} ({0} bytes)'
                      .format(self.native_data_size))
        # I thought this might be an OLE type specifier ???
        self.unknown_short, index = read_uint16(data, index)
        self.filename, index = read_zero_terminated_string(data, index)
        # source path
        self.src_path, index = read_zero_terminated_string(data, index)
        # TODO: I bet these 8 bytes are a timestamp ==> FILETIME from olefile
        self.unknown_long_1, index = read_uint32(data, index)
        self.unknown_long_2, index = read_uint32(data, index)
        # temp path?
        self.temp_path, index = read_zero_terminated_string(data, index)
        # size of the rest of the data
        try:
            self.actual_size, index = read_uint32(data, index)
            if self.data_is_stream:
                self.data = data
            else:
                self.data = data[index:index+self.actual_size]
            self.is_link = False
            # TODO: there can be extra data, no idea what it is for
            # TODO: SLACK DATA
        except (IOError, struct.error):      # no data to read actual_size
            log.debug('data is not embedded but only a link')
            self.is_link = True
            self.actual_size = 0
            self.data = None


class OleObject(object):
    """
    OLE 1.0 Object

    see MS-OLEDS 2.2 OLE1.0 Format Structures
    """

    # constants for the format_id attribute:
    # see MS-OLEDS 2.2.4 ObjectHeader
    TYPE_LINKED = 0x01
    TYPE_EMBEDDED = 0x02

    def __init__(self, bindata=None):
        """
        Constructor for OleObject.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: bytes, OLE 1.0 Object structure containing OLE object

        Note: Code can easily by generalized to work with byte streams instead
              of arrays just like in OleNativeStream.
        """
        self.ole_version = None
        self.format_id = None
        self.class_name = None
        self.topic_name = None
        self.item_name = None
        self.data = None
        self.data_size = None
        if bindata is not None:
            self.parse(bindata)

    def parse(self, data):
        """
        Parse binary data containing an OLE 1.0 Object structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.2 OLE1.0 Format Structures)

        :param data: bytes, OLE 1.0 Object structure containing an OLE object
        :return:
        """
        # from ezhexviewer import hexdump3
        # print("Parsing OLE object data:")
        # print(hexdump3(data, length=16))
        # Header: see MS-OLEDS 2.2.4 ObjectHeader
        index = 0
        self.ole_version, index = read_uint32(data, index)
        self.format_id, index = read_uint32(data, index)
        log.debug('OLE version=%08X - Format ID=%08X',
                  self.ole_version, self.format_id)
        assert self.format_id in (self.TYPE_EMBEDDED, self.TYPE_LINKED)
        self.class_name, index = read_length_prefixed_string(data, index)
        self.topic_name, index = read_length_prefixed_string(data, index)
        self.item_name, index = read_length_prefixed_string(data, index)
        log.debug('Class name=%r - Topic name=%r - Item name=%r',
                  self.class_name, self.topic_name, self.item_name)
        if self.format_id == self.TYPE_EMBEDDED:
            # Embedded object: see MS-OLEDS 2.2.5 EmbeddedObject
            # assert self.topic_name != '' and self.item_name != ''
            self.data_size, index = read_uint32(data, index)
            log.debug('Declared data size=%d - remaining size=%d',
                      self.data_size, len(data)-index)
            # TODO: handle incorrect size to avoid exception
            self.data = data[index:index+self.data_size]
            assert len(self.data) == self.data_size
            self.extra_data = data[index+self.data_size:]


def shorten_filename(fname, max_len):
    """Create filename shorter than max_len, trying to preserve suffix."""
    # simple cases:
    if not max_len:
        return fname
    name_len = len(fname)
    if name_len < max_len:
        return fname

    idx = fname.rfind('.')
    if idx == -1:
        return fname[:max_len]

    suffix_len = name_len - idx  # length of suffix including '.'
    if suffix_len > max_len:
        return fname[:max_len]

    # great, can preserve suffix
    return fname[:max_len-suffix_len] + fname[idx:]


def sanitize_filename(filename, replacement='_',
                      max_len=MAX_FILENAME_LENGTH):
    """
    Return filename that is save to work with.

    Removes path components, replaces all non-whitelisted characters (so output
    is always a pure-ascii string), replaces '..' and '  ' and shortens to
    given max length, trying to preserve suffix.

    Might return empty string
    """
    basepath = os.path.basename(filename).strip()
    sane_fname = re.sub(u'[^a-zA-Z0-9._ -]', replacement, basepath)
    sane_fname = str(sane_fname)    # py3: does nothing;   py2: unicode --> str

    while ".." in sane_fname:
        sane_fname = sane_fname.replace('..', '.')

    while "  " in sane_fname:
        sane_fname = sane_fname.replace('  ', ' ')

    # limit filename length, try to preserve suffix
    return shorten_filename(sane_fname, max_len)


def get_sane_embedded_filenames(filename, src_path, tmp_path, max_len,
                                noname_index):
    """
    Get some sane filenames out of path information, preserving file suffix.

    Returns several canddiates, first with suffix, then without, then random
    with suffix and finally one last attempt ignoring max_len using arg
    `noname_index`.

    In some malware examples, filename (on which we relied sofar exclusively
    for this) is empty or " ", but src_path and tmp_path contain paths with
    proper file names. Try to extract filename from any of those.

    Preservation of suffix is especially important since that controls how
    windoze treats the file.
    """
    suffixes = []
    candidates_without_suffix = []  # remember these as fallback
    for candidate in (filename, src_path, tmp_path):
        # remove path component. Could be from linux, mac or windows
        idx = max(candidate.rfind('/'), candidate.rfind('\\'))
        candidate = candidate[idx+1:].strip()

        # sanitize
        candidate = sanitize_filename(candidate, max_len=max_len)

        if not candidate:
            continue    # skip whitespace-only

        # identify suffix. Dangerous suffixes are all short
        idx = candidate.rfind('.')
        if idx == -1:
            candidates_without_suffix.append(candidate)
            continue
        elif idx < len(candidate)-5:
            candidates_without_suffix.append(candidate)
            continue

        # remember suffix
        suffixes.append(candidate[idx:])

        yield candidate

    # parts with suffix not good enough? try those without one
    for candidate in candidates_without_suffix:
        yield candidate

    # then try random
    suffixes.append('')  # ensure there is something in there
    for _ in range(MAX_FILENAME_ATTEMPTS):
        for suffix in suffixes:
            leftover_len = max_len - len(suffix)
            if leftover_len < 1:
                continue
            name = ''.join(random.sample('abcdefghijklmnopqrstuvwxyz',
                                         min(26, leftover_len)))
            yield name + suffix

    # still not returned? Then we have to make up a name ourselves
    # do not care any more about max_len (maybe it was 0 or negative)
    yield 'oleobj_%03d' % noname_index


def find_ole_in_ppt(filename):
    """ find ole streams in ppt

    This may be a bit confusing: we get an ole file (or its name) as input and
    as output we produce possibly several ole files. This is because the
    data structure can be pretty nested:
    A ppt file has many streams that consist of records. Some of these records
    can contain data which contains data for another complete ole file (which
    we yield). This embedded ole file can have several streams, one of which
    can contain the actual embedded file we are looking for (caller will check
    for these).
    """
    ppt_file = None
    try:
        ppt_file = PptFile(filename)
        for stream in ppt_file.iter_streams():
            for record_idx, record in enumerate(stream.iter_records()):
                if isinstance(record, PptRecordExOleVbaActiveXAtom):
                    ole = None
                    try:
                        data_start = next(record.iter_uncompressed())
                        if data_start[:len(olefile.MAGIC)] != olefile.MAGIC:
                            continue   # could be ActiveX control / VBA Storage

                        # otherwise, this should be an OLE object
                        log.debug('Found record with embedded ole object in '
                                  'ppt (stream "{0}", record no {1})'
                                  .format(stream.name, record_idx))
                        ole = record.get_data_as_olefile()
                        yield ole
                    except IOError:
                        log.warning('Error reading data from {0} stream or '
                                    'interpreting it as OLE object'
                                    .format(stream.name))
                        log.debug('', exc_info=True)
                    finally:
                        if ole is not None:
                            ole.close()
    finally:
        if ppt_file is not None:
            ppt_file.close()


class FakeFile(io.RawIOBase):
    """ create file-like object from data without copying it

    BytesIO is what I would like to use but it copies all the data. This class
    does not. On the downside: data can only be read and seeked, not written.

    Assume that given data is bytes (str in py2, bytes in py3).

    See also (and maybe can put into common file with):
    ppt_record_parser.IterStream, ooxml.ZipSubFile
    """

    def __init__(self, data):
        """ create FakeFile with given bytes data """
        super(FakeFile, self).__init__()
        self.data = data   # this does not actually copy (python is lazy)
        self.pos = 0
        self.size = len(data)

    def readable(self):
        return True

    def writable(self):
        return False

    def seekable(self):
        return True

    def readinto(self, target):
        """ read into pre-allocated target """
        n_data = min(len(target), self.size-self.pos)
        if n_data == 0:
            return 0
        target[:n_data] = self.data[self.pos:self.pos+n_data]
        self.pos += n_data
        return n_data

    def read(self, n_data=-1):
        """ read and return data """
        if self.pos >= self.size:
            return bytes()
        if n_data == -1:
            n_data = self.size - self.pos
        result = self.data[self.pos:self.pos+n_data]
        self.pos += n_data
        return result

    def seek(self, pos, offset=io.SEEK_SET):
        """ jump to another position in file """
        # calc target position from self.pos, pos and offset
        if offset == io.SEEK_SET:
            new_pos = pos
        elif offset == io.SEEK_CUR:
            new_pos = self.pos + pos
        elif offset == io.SEEK_END:
            new_pos = self.size + pos
        else:
            raise ValueError("invalid offset {0}, need SEEK_* constant"
                             .format(offset))
        if new_pos < 0:
            raise IOError('Seek beyond start of file not allowed')
        self.pos = new_pos

    def tell(self):
        """ tell where in file we are positioned """
        return self.pos


def find_ole(filename, data, xml_parser=None):
    """ try to open somehow as zip/ole/rtf/... ; yield None if fail

    If data is given, filename is (mostly) ignored.

    yields embedded ole streams in form of OleFileIO.
    """

    if data is not None:
        # isOleFile and is_ppt can work on data directly but zip need file
        # --> wrap data in a file-like object without copying data
        log.debug('working on data, file is not touched below')
        arg_for_ole = data
        arg_for_zip = FakeFile(data)
    else:
        # we only have a file name
        log.debug('working on file by name')
        arg_for_ole = filename
        arg_for_zip = filename

    ole = None
    try:
        if olefile.isOleFile(arg_for_ole):
            if is_ppt(arg_for_ole):
                log.info('is ppt file: ' + filename)
                for ole in find_ole_in_ppt(arg_for_ole):
                    yield ole
                    ole = None   # is closed in find_ole_in_ppt
            # in any case: check for embedded stuff in non-sectored streams
            log.info('is ole file: ' + filename)
            ole = olefile.OleFileIO(arg_for_ole)
            yield ole
        elif xml_parser is not None or is_zipfile(arg_for_zip):
            # keep compatibility with 3rd-party code that calls this function
            # directly without providing an XmlParser instance
            if xml_parser is None:
                xml_parser = XmlParser(arg_for_zip)
                # force iteration so XmlParser.iter_non_xml() returns data
                for _ in xml_parser.iter_xml():
                    pass

            log.info('is zip file: ' + filename)
            # we looped through the XML files before, now we can
            # iterate the non-XML files looking for ole objects
            for subfile, _, file_handle in xml_parser.iter_non_xml():
                try:
                    head = file_handle.read(len(olefile.MAGIC))
                except RuntimeError:
                    log.error('zip is encrypted: ' + filename)
                    yield None
                    continue

                if head == olefile.MAGIC:
                    file_handle.seek(0)
                    log.info('  unzipping ole: ' + subfile)
                    try:
                        ole = olefile.OleFileIO(file_handle)
                        yield ole
                    except IOError:
                        log.warning('Error reading data from {0}/{1} or '
                                    'interpreting it as OLE object'
                                    .format(filename, subfile))
                        log.debug('', exc_info=True)
                    finally:
                        if ole is not None:
                            ole.close()
                            ole = None
                else:
                    log.debug('unzip skip: ' + subfile)
        else:
            log.warning('open failed: {0} (or its data) is neither zip nor OLE'
                        .format(filename))
            yield None
    except Exception:
        log.error('Caught exception opening {0}'.format(filename),
                  exc_info=True)
        yield None
    finally:
        if ole is not None:
            ole.close()


def find_external_relationships(xml_parser):
    """ iterate XML files looking for relationships to external objects
    """
    for _, elem, _ in xml_parser.iter_xml(None, False, OOXML_RELATIONSHIP_TAG):
        try:
            if elem.attrib['TargetMode'] == 'External':
                relationship_type = elem.attrib['Type'].rsplit('/', 1)[1]

                if relationship_type in BLACKLISTED_RELATIONSHIP_TYPES:
                    yield relationship_type, elem.attrib['Target']
        except (AttributeError, KeyError):
            # ignore missing attributes - Word won't detect
            # external links anyway
            pass


def find_customUI(xml_parser):
    """
    iterate XML files looking for customUI to external objects or VBA macros
    Examples of malicious usage, to load an external document or trigger a VBA macro:
    https://www.trellix.com/en-us/about/newsroom/stories/threat-labs/prime-ministers-office-compromised.html
    https://www.netero1010-securitylab.com/evasion/execution-of-remote-vba-script-in-excel
    """
    for _, elem, _ in xml_parser.iter_xml(None, False, (TAG_CUSTOMUI_2007, TAG_CUSTOMUI_2010)):
       customui_onload = elem.get('onLoad')
       if customui_onload is not None:
            yield customui_onload


def process_file(filename, data, output_dir=None):
    """ find embedded objects in given file

    if data is given (from xglob for encrypted zip files), then filename is
    not used for reading. If not (usual case), then data is read from filename
    on demand.

    If output_dir is given and does not exist, it is created. If it is not
    given, data is saved to same directory as the input file.
    """
    # sanitize filename, leave space for embedded filename part
    sane_fname = sanitize_filename(filename, max_len=MAX_FILENAME_LENGTH-5) or\
        'NONAME'
    if output_dir:
        if not os.path.isdir(output_dir):
            log.info('creating output directory %s', output_dir)
            os.mkdir(output_dir)

        fname_prefix = os.path.join(output_dir, sane_fname)
    else:
        base_dir = os.path.dirname(filename)
        fname_prefix = os.path.join(base_dir, sane_fname)

    # TODO: option to extract objects to files (false by default)
    print('-'*79)
    print('File: %r' % filename)
    index = 1

    # do not throw errors but remember them and try continue with other streams
    err_stream = False
    err_dumping = False
    did_dump = False

    xml_parser = None
    if is_zipfile(filename):
        log.info('file could be an OOXML file, looking for relationships with '
                 'external links')
        xml_parser = XmlParser(filename)
        for relationship, target in find_external_relationships(xml_parser):
            did_dump = True
            print("Found relationship '%s' with external link %s" % (relationship, target))
            if target.startswith('mhtml:'):
                print("Potential exploit for CVE-2021-40444")
        for target in find_customUI(xml_parser):
            did_dump = True
            print("Found customUI tag with external link or VBA macro %s (possibly exploiting CVE-2021-42292)" % target)

    # look for ole files inside file (e.g. unzip docx)
    # have to finish work on every ole stream inside iteration, since handles
    # are closed in find_ole
    for ole in find_ole(filename, data, xml_parser):
        if ole is None:    # no ole file found
            continue

        for path_parts in ole.listdir():
            stream_path = '/'.join(path_parts)
            log.debug('Checking stream %r', stream_path)
            if path_parts[-1].lower() == '\x01ole10native':
                stream = None
                try:
                    stream = ole.openstream(path_parts)
                    print('extract file embedded in OLE object from stream %r:'
                          % stream_path)
                    print('Parsing OLE Package')
                    opkg = OleNativeStream(stream)
                    # leave stream open until dumping is finished
                except Exception:
                    log.warning('*** Not an OLE 1.0 Object')
                    err_stream = True
                    if stream is not None:
                        stream.close()
                    continue

                # print info
                if opkg.is_link:
                    log.debug('Object is not embedded but only linked to '
                              '- skip')
                    continue
                print(u'Filename = "%s"' % opkg.filename)
                print(u'Source path = "%s"' % opkg.src_path)
                print(u'Temp path = "%s"' % opkg.temp_path)
                for embedded_fname in get_sane_embedded_filenames(
                        opkg.filename, opkg.src_path, opkg.temp_path,
                        MAX_FILENAME_LENGTH - len(sane_fname) - 1, index):
                    fname = fname_prefix + '_' + embedded_fname
                    if not os.path.isfile(fname):
                        break

                # dump
                try:
                    print('saving to file %s' % fname)
                    with open(fname, 'wb') as writer:
                        n_dumped = 0
                        next_size = min(DUMP_CHUNK_SIZE, opkg.actual_size)
                        while next_size:
                            data = stream.read(next_size)
                            writer.write(data)
                            n_dumped += len(data)
                            if len(data) != next_size:
                                log.warning('Wanted to read {0}, got {1}'
                                            .format(next_size, len(data)))
                                break
                            next_size = min(DUMP_CHUNK_SIZE,
                                            opkg.actual_size - n_dumped)
                    did_dump = True
                except Exception as exc:
                    log.warning('error dumping to {0} ({1})'
                                .format(fname, exc))
                    err_dumping = True
                finally:
                    stream.close()

                index += 1
    return err_stream, err_dumping, did_dump


# === MAIN ====================================================================


def existing_file(filename):
    """ called by argument parser to see whether given file exists """
    if not os.path.isfile(filename):
        raise argparse.ArgumentTypeError('{0} is not a file.'.format(filename))
    return filename


def main(cmd_line_args=None):
    """ main function, called when running this as script

    Per default (cmd_line_args=None) uses sys.argv. For testing, however, can
    provide other arguments.
    """
    # print banner with version
    ensure_stdout_handles_unicode()
    print('oleobj %s - http://decalage.info/oletools' % __version__)
    print('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print('Please report any issue at '
          'https://github.com/decalage2/oletools/issues')
    print('')

    usage = 'usage: %(prog)s [options] <filename> [filename2 ...]'
    parser = argparse.ArgumentParser(usage=usage)
    # parser.add_argument('-o', '--outfile', dest='outfile',
    #     help='output file')
    # parser.add_argument('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_argument("-r", action="store_true", dest="recursive",
                        help='find files recursively in subdirectories.')
    parser.add_argument("-d", type=str, dest="output_dir", default=None,
                        help='use specified directory to output files.')
    parser.add_argument("-z", "--zip", dest='zip_password', type=str,
                        default=None,
                        help='if the file is a zip archive, open first file '
                             'from it, using the provided password (requires '
                             'Python 2.6+)')
    parser.add_argument("-f", "--zipfname", dest='zip_fname', type=str,
                        default='*',
                        help='if the file is a zip archive, file(s) to be '
                             'opened within the zip. Wildcards * and ? are '
                             'supported. (default:*)')
    parser.add_argument('-l', '--loglevel', dest="loglevel", action="store",
                        default=DEFAULT_LOG_LEVEL,
                        help='logging level debug/info/warning/error/critical '
                             '(default=%(default)s)')
    parser.add_argument('input', nargs='*', type=existing_file, metavar='FILE',
                        help='Office files to parse (same as -i)')

    # options for compatibility with ripOLE
    parser.add_argument('-i', '--more-input', type=str, metavar='FILE',
                        help='Additional file to parse (same as positional '
                             'arguments)')
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='verbose mode, set logging to DEBUG '
                             '(overwrites -l)')

    options = parser.parse_args(cmd_line_args)
    if options.more_input:
        options.input += [options.more_input, ]
    if options.verbose:
        options.loglevel = 'debug'

    # Print help if no arguments are passed
    if not options.input:
        parser.print_help()
        return RETURN_ERR_ARGS

    # Setup logging to the console:
    # here we use stdout instead of stderr by default, so that the output
    # can be redirected properly.
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)
    if options.loglevel == 'debug-olefile':
        olefile.enable_logging()

    # remember if there was a problem and continue with other data
    any_err_stream = False
    any_err_dumping = False
    any_did_dump = False

    for container, filename, data in \
            xglob.iter_files(options.input, recursive=options.recursive,
                             zip_password=options.zip_password,
                             zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        err_stream, err_dumping, did_dump = \
            process_file(filename, data, options.output_dir)
        any_err_stream |= err_stream
        any_err_dumping |= err_dumping
        any_did_dump |= did_dump

    # assemble return value
    return_val = RETURN_NO_DUMP
    if any_did_dump:
        return_val += RETURN_DID_DUMP
    if any_err_stream:
        return_val += RETURN_ERR_STREAM
    if any_err_dumping:
        return_val += RETURN_ERR_DUMP
    return return_val


if __name__ == '__main__':
    sys.exit(main())
