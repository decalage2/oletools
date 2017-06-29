#!/usr/bin/env python
from __future__ import print_function
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

# oleobj is copyright (c) 2015-2017 Philippe Lagadec (http://www.decalage.info)
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
# 2016-06          PL: - added main and process_file (not working yet)
# 2016-07-18 v0.48 SL: - added Python 3.5 support
# 2016-07-19       PL: - fixed Python 2.6-7 support
# 2016-11-17 v0.51 PL: - fixed OLE native object extraction
# 2016-11-18       PL: - added main for setup.py entry point
# 2017-05-03       PL: - fixed absolute imports (issue #141)

__version__ = '0.51'

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

import logging, struct, optparse, os, re, sys

# IMPORTANT: it should be possible to run oletools directly as scripts
# in any directory without installing them with pip or setup.py.
# In that case, relative imports are NOT usable.
# And to enable Python 2+3 compatibility, we need to use absolute imports,
# so we add the oletools parent folder to sys.path (absolute+normalized path):
_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if not _parent_dir in sys.path:
    sys.path.insert(0, _parent_dir)

from oletools.thirdparty.olefile import olefile
from oletools.thirdparty.xglob import xglob

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

def enable_logging():
    """
    Enable logging for this module (disabled by default).
    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    log.setLevel(logging.NOTSET)


# === CONSTANTS ==============================================================

# some str methods on Python 2.x return characters,
# while the equivalent bytes methods return integers on Python 3.x:
if sys.version_info[0] <= 2:
    # Python 2.x
    NULL_CHAR = '\x00'
else:
    # Python 3.x
    NULL_CHAR = 0


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
    assert data[length] == NULL_CHAR
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


    def __init__(self, bindata=None, package=False):
        """
        Constructor for OleNativeStream.
        If bindata is provided, it will be parsed using the parse() method.

        :param bindata: bytes, OLENativeStream structure containing an OLE object
        :param package: bool, set to True when extracting from an OLE Package object
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
        # An OLE Package object does not have the native data size field
        if not self.package:
            self.native_data_size = struct.unpack('<L', data[0:4])[0]
            data = data[4:]
            log.debug('OLE native data size = {0:08X} ({0} bytes)'.format(self.native_data_size))
        # I thought this might be an OLE type specifier ???
        self.unknown_short, data = read_uint16(data)
        self.filename, data = data.split(b'\x00', 1)
        # source path
        self.src_path, data = data.split(b'\x00', 1)
        # TODO I bet these next 8 bytes are a timestamp => FILETIME from olefile
        self.unknown_long_1, data = read_uint32(data)
        self.unknown_long_2, data = read_uint32(data)
        # temp path?
        self.temp_path, data = data.split(b'\x00', 1)
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
        # from ezhexviewer import hexdump3
        # print("Parsing OLE object data:")
        # print(hexdump3(data, length=16))
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



def sanitize_filename(filename, replacement='_', max_length=200):
    """compute basename of filename. Replaces all non-whitelisted characters.
       The returned filename is always a basename of the file."""
    basepath = os.path.basename(filename).strip()
    sane_fname = re.sub(r'[^\w\.\- ]', replacement, basepath)

    while ".." in sane_fname:
        sane_fname = sane_fname.replace('..', '.')

    while "  " in sane_fname:
        sane_fname = sane_fname.replace('  ', ' ')

    if not len(filename):
        sane_fname = 'NONAME'

    # limit filename length
    if max_length:
        sane_fname = sane_fname[:max_length]

    return sane_fname


def process_file(container, filename, data, output_dir=None):
    if output_dir:
        if not os.path.isdir(output_dir):
            log.info('creating output directory %s' % output_dir)
            os.mkdir(output_dir)

        fname_prefix = os.path.join(output_dir,
                                    sanitize_filename(filename))
    else:
        base_dir = os.path.dirname(filename)
        sane_fname = sanitize_filename(filename)
        fname_prefix = os.path.join(base_dir, sane_fname)

    # TODO: option to extract objects to files (false by default)
    if data is None:
        data = open(filename, 'rb').read()
    print ('-'*79)
    print ('File: %r - %d bytes' % (filename, len(data)))
    ole = olefile.OleFileIO(data)
    index = 1
    for stream in ole.listdir():
        if stream[-1] == '\x01Ole10Native':
            objdata = ole.openstream(stream).read()
            stream_path = '/'.join(stream)
            log.debug('Checking stream %r' % stream_path)
            try:
                print('extract file embedded in OLE object from stream %r:' % stream_path)
                print ('Parsing OLE Package')
                opkg = OleNativeStream(bindata=objdata)
                print ('Filename = %r' % opkg.filename)
                print ('Source path = %r' % opkg.src_path)
                print ('Temp path = %r' % opkg.temp_path)
                if opkg.filename:
                    fname = '%s_%s' % (fname_prefix,
                                       sanitize_filename(opkg.filename))
                else:
                    fname = '%s_object_%03d.noname' % (fname_prefix, index)
                print ('saving to file %s' % fname)
                open(fname, 'wb').write(opkg.data)
                index += 1
            except:
                log.debug('*** Not an OLE 1.0 Object')



#=== MAIN =================================================================

def main():
    # print banner with version
    print ('oleobj %s - http://decalage.info/oletools' % __version__)
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')

    DEFAULT_LOG_LEVEL = "warning" # Default log level
    LOG_LEVELS = {'debug':    logging.DEBUG,
              'info':     logging.INFO,
              'warning':  logging.WARNING,
              'error':    logging.ERROR,
              'critical': logging.CRITICAL
             }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    # parser.add_option('-o', '--outfile', dest='outfile',
    #     help='output file')
    # parser.add_option('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_option("-r", action="store_true", dest="recursive",
        help='find files recursively in subdirectories.')
    parser.add_option("-d", type="str", dest="output_dir",
        help='use specified directory to output files.', default=None)
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
        help='if the file is a zip archive, open first file from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
        help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print (__doc__)
        parser.print_help()
        sys.exit()

    # Setup logging to the console:
    # here we use stdout instead of stderr by default, so that the output
    # can be redirected properly.
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)


    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
        zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        process_file(container, filename, data, options.output_dir)

if __name__ == '__main__':
    main()

