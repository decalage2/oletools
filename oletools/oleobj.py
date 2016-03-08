#!/usr/bin/env python
"""
oleobj.py

oleobj is a Python script and module to parse OLE objects and files stored
into various file formats such as RTF or MS Office documents (e.g. Word, Excel).

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

oleobj is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# oleobj is copyright (c) 2015 Philippe Lagadec (http://www.decalage.info)
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
# 2015-12-05 v0.01 PL: - first version

__version__ = '0.01'

#------------------------------------------------------------------------------
# TODO:
# + setup logging (common with other oletools)


#------------------------------------------------------------------------------
# REFERENCES:

# Reference for the storage of embedded OLE objects/files:
# [MS-OLEDS]: Object Linking and Embedding (OLE) Data Structures
# https://msdn.microsoft.com/en-us/library/dd942265.aspx

# - officeparser: https://github.com/unixfreak0037/officeparser
# TODO: oledump


#--- IMPORTS ------------------------------------------------------------------

import logging, struct


# === LOGGING =================================================================

class NullHandler(logging.Handler):
    """
    Log Handler without output, to avoid printing messages if logging is not
    configured by the main application.
    Python 2.7 has logging.NullHandler, but this is necessary for 2.6:
    see https://docs.python.org/2.6/library/logging.html#configuring-logging-for-a-library
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
        #NOTE: another less intrusive but more "hackish" solution would be to
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
log = get_logger('oleobj')


# === GLOBAL VARIABLES =======================================================

# struct to parse an unsigned integer of 32 bits:
struct_uint32 = struct.Struct('<L')
assert struct_uint32.size == 4  # make sure it matches 4 bytes

# struct to parse an unsigned integer of 16 bits:
struct_uint16 = struct.Struct('<H')
assert struct_uint16.size == 2  # make sure it matches 2 bytes


# === FUNCTIONS ==============================================================

def read_uint32(data):
    """
    Read an unsigned integer from the first 32 bits of data.

    :param data: bytes string containing the data to be extracted.
    :return: tuple (value, new_data) containing the read value (int),
             and the new data without the bytes read.
    """
    value = struct_uint32.unpack(data[0:4])[0]
    new_data = data[4:]
    return (value, new_data)


def read_uint16(data):
    """
    Read an unsigned integer from the first 16 bits of data.

    :param data: bytes string containing the data to be extracted.
    :return: tuple (value, new_data) containing the read value (int),
             and the new data without the bytes read.
    """
    value = struct_uint16.unpack(data[0:2])[0]
    new_data = data[2:]
    return (value, new_data)


def read_LengthPrefixedAnsiString(data):
    """
    Read a length-prefixed ANSI string from data.

    :param data: bytes string containing the data to be extracted.
    :return: tuple (value, new_data) containing the read value (bytes string),
             and the new data without the bytes read.
    """
    length, data = read_uint32(data)
    # if length = 0, return a null string (no null character)
    if length == 0:
        return ('', data)
    # extract the string without the last null character
    ansi_string = data[:length-1]
    # TODO: only in strict mode:
    # check the presence of the null char:
    assert data[length] == '\x00'
    new_data = data[length:]
    return (ansi_string, new_data)


# === CLASSES ================================================================

class OleNativeStream (object):
    """
    OLE object contained into an OLENativeStream structure.
    (see MS-OLEDS 2.3.6 OLENativeStream)
    """
    # constants for the type attribute:
    # see MS-OLEDS 2.2.4 ObjectHeader
    TYPE_LINKED = 0x01
    TYPE_EMBEDDED = 0x02


    def __init__(self, bindata=None):
        """
        Constructor for OleNativeStream.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: bytes, OLENativeStream structure containing an OLE object
        """
        self.filename = None
        self.src_path = None
        self.unknown_short = None
        self.unknown_long_1 = None
        self.unknown_long_2 = None
        self.temp_path = None
        self.actual_size = None
        self.data = None
        if bindata is not None:
            self.parse(data=bindata)

    def parse(self, data):
        """
        Parse binary data containing an OLENativeStream structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.3.6 OLENativeStream)

        :param data: bytes, OLENativeStream structure containing an OLE object
        :return:
        """
        # TODO: strict mode to raise exceptions when values are incorrect
        # (permissive mode by default)
        # self.native_data_size = struct.unpack('<L', data[0:4])[0]
        # data = data[4:]
        # log.debug('OLE native data size = {0:08X} ({0} bytes)'.format(self.native_data_size))
        # I thought this might be an OLE type specifier ???
        self.unknown_short, data = read_uint16(data)
        self.filename, data = data.split('\x00', 1)
        # source path
        self.src_path, data = data.split('\x00', 1)
        # TODO I bet these next 8 bytes are a timestamp => FILETIME from olefile
        self.unknown_long_1, data = read_uint32(data)
        self.unknown_long_2, data = read_uint32(data)
        # temp path?
        self.temp_path, data = data.split('\x00', 1)
        # size of the rest of the data
        self.actual_size, data = read_uint32(data)
        self.data = data[0:self.actual_size]
        # TODO: exception when size > remaining data
        # TODO: SLACK DATA


class OleObject (object):
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

        :param bindata: bytes, OLE 1.0 Object structure containing an OLE object
        """
        self.ole_version = None
        self.format_id = None
        self.class_name = None
        self.topic_name = None
        self.item_name = None
        self.data = None
        self.data_size = None

    def parse(self, data):
        """
        Parse binary data containing an OLE 1.0 Object structure,
        to extract the OLE object it contains.
        (see MS-OLEDS 2.2 OLE1.0 Format Structures)

        :param data: bytes, OLE 1.0 Object structure containing an OLE object
        :return:
        """
        # Header: see MS-OLEDS 2.2.4 ObjectHeader
        self.ole_version, data = read_uint32(data)
        self.format_id, data = read_uint32(data)
        log.debug('OLE version=%08X - Format ID=%08X' % (self.ole_version, self.format_id))
        assert self.format_id in (self.TYPE_EMBEDDED, self.TYPE_LINKED)
        self.class_name, data = read_LengthPrefixedAnsiString(data)
        self.topic_name, data = read_LengthPrefixedAnsiString(data)
        self.item_name, data = read_LengthPrefixedAnsiString(data)
        log.debug('Class name=%r - Topic name=%r - Item name=%r'
                      % (self.class_name, self.topic_name, self.item_name))
        if self.format_id == self.TYPE_EMBEDDED:
            # Embedded object: see MS-OLEDS 2.2.5 EmbeddedObject
            #assert self.topic_name != '' and self.item_name != ''
            self.data_size, data = read_uint32(data)
            log.debug('Declared data size=%d - remaining size=%d' % (self.data_size, len(data)))
            # TODO: handle incorrect size to avoid exception
            self.data = data[:self.data_size]
            assert len(self.data) == self.data_size
            self.extra_data = data[self.data_size:]
