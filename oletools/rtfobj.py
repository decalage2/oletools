#!/usr/bin/env python
"""
rtfobj.py

rtfobj is a Python module to extract embedded objects from RTF files, such as
OLE ojects. It can be used as a Python library or a command-line tool.

Usage: rtfobj.py <file.rtf>

rtfobj project website: http://www.decalage.info/python/rtfobj

rtfobj is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# rtfobj is copyright (c) 2012-2016, Philippe Lagadec (http://www.decalage.info)
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
# 2012-11-09 v0.01 PL: - first version
# 2013-04-02 v0.02 PL: - fixed bug in main
# 2015-12-09 v0.03 PL: - configurable logging, CLI options
#                      - extract OLE 1.0 objects
#                      - extract files from OLE Package objects
# 2016-04-01 v0.04 PL: - fixed logging output to use stdout instead of stderr
# 2016-04-07 v0.45 PL: - improved parsing to handle some malware tricks
# 2016-05-06 v0.47 TJ: - added option -d to set the output directory
#                        (contribution by Thomas Jarosch)
#                  TJ: - sanitize filenames to avoid special characters
# 2016-05-29       PL: - improved parsing, fixed issue #42

__version__ = '0.47'

#------------------------------------------------------------------------------
# TODO:
# - improve regex pattern for better performance?
# - allow semicolon within hex, as found in  this sample:
#   http://contagiodump.blogspot.nl/2011/10/sep-28-cve-2010-3333-manuscript-with.html


#=== IMPORTS =================================================================

import re, os, sys, string, binascii, logging, optparse

from thirdparty.xglob import xglob
from oleobj import OleObject, OleNativeStream
import oleobj

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
log = get_logger('rtfobj')


#=== CONSTANTS=================================================================

# REGEX pattern to extract embedded OLE objects in hexadecimal format:

# alphanum digit: [0-9A-Fa-f]
HEX_DIGIT = r'[0-9A-Fa-f]'

# hex char = two alphanum digits: [0-9A-Fa-f]{2}
# HEX_CHAR = r'[0-9A-Fa-f]{2}'
# in fact MS Word allows whitespaces in between the hex digits!
# HEX_CHAR = r'[0-9A-Fa-f]\s*[0-9A-Fa-f]'
# Even worse, MS Word also allows ANY RTF-style tag {*} in between!!
# AND the tags can be nested...
SINGLE_RTF_TAG = r'[{][^{}]*[}]'
# Nested tags, two levels (because Python's re does not support nested matching):
NESTED_RTF_TAG = r'[{](?:[^{}]|'+SINGLE_RTF_TAG+r')*[}]'
# ignored whitespaces and tags within a hex block:
IGNORED = r'(?:\s|'+NESTED_RTF_TAG+r')*'
#IGNORED = r'\s*'

# HEX_CHAR = HEX_DIGIT + IGNORED + HEX_DIGIT

# several hex chars, at least 4: (?:[0-9A-Fa-f]{2}){4,}
# + word boundaries
# HEX_CHARS_4orMORE = r'\b(?:' + HEX_CHAR + r'){4,}\b'
# at least 1 hex char:
# HEX_CHARS_1orMORE = r'(?:' + HEX_CHAR + r')+'
# at least 1 hex char, followed by whitespace or CR/LF:
# HEX_CHARS_1orMORE_WHITESPACES = r'(?:' + HEX_CHAR + r')+\s+'
# + word boundaries around hex block
# HEX_CHARS_1orMORE_WHITESPACES = r'\b(?:' + HEX_CHAR + r')+\b\s*'
# at least one block of hex and whitespace chars, followed by closing curly bracket:
# HEX_BLOCK_CURLY_BRACKET = r'(?:' + HEX_CHARS_1orMORE_WHITESPACES + r')+\}'
# PATTERN = r'(?:' + HEX_CHARS_1orMORE_WHITESPACES + r')*' + HEX_CHARS_1orMORE

#TODO PATTERN = r'\b(?:' + HEX_CHAR + IGNORED + r'){4,}\b'
# PATTERN = r'\b(?:' + HEX_CHAR + IGNORED + r'){4,}' #+ HEX_CHAR + r'\b'
PATTERN = r'\b(?:' + HEX_DIGIT + IGNORED + r'){7,}' + HEX_DIGIT + r'\b'

# at least 4 hex chars, followed by whitespace or CR/LF: (?:[0-9A-Fa-f]{2}){4,}\s*
# PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'
# improved pattern, allowing semicolons within hex:
#PATTERN = r'(?:(?:[0-9A-Fa-f]{2})+\s*)*(?:[0-9A-Fa-f]{2}){4,}'

# a dummy translation table for str.translate, which does not change anythying:
TRANSTABLE_NOCHANGE = string.maketrans('', '')

re_hexblock = re.compile(PATTERN)
re_embedded_tags = re.compile(IGNORED)
re_decimal = re.compile(r'\d+')

re_delimiter = re.compile(r'[ \t\r\n\f\v]')

DELIMITER = r'[ \t\r\n\f\v]'
DELIMITERS_ZeroOrMore = r'[ \t\r\n\f\v]*'
BACKSLASH_BIN = r'\\bin'
# According to my tests, Word accepts up to 250 digits (leading zeroes)
DECIMAL_GROUP = r'(\d{1,250})'

re_delims_bin_decimal = re.compile(DELIMITERS_ZeroOrMore + BACKSLASH_BIN
                                   + DECIMAL_GROUP + DELIMITER)
re_delim_hexblock = re.compile(DELIMITER + PATTERN)


#=== FUNCTIONS ===============================================================

def rtf_iter_objects_old (filename, min_size=32):
    """
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    data = open(filename, 'rb').read()
    for m in re.finditer(PATTERN, data):
        found = m.group(0)
        orig_len = len(found)
        # remove all whitespace and line feeds:
        #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
        found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v}')
        found = binascii.unhexlify(found)
        #print repr(found)
        if len(found)>min_size:
            yield m.start(), orig_len, found

# TODO: backward-compatible API?


def search_hex_block(data, pos=0, min_size=32, first=True):
    if first:
        # Search 1st occurence of a hex block:
        match = re_hexblock.search(data, pos=pos)
    else:
        # Match next occurences of a hex block, from the current position only:
        match = re_hexblock.match(data, pos=pos)



def rtf_iter_objects (data, min_size=32):
    """
    Open a RTF file, extract each embedded object encoded in hexadecimal of
    size > min_size, yield the index of the object in the RTF file and its data
    in binary format.
    This is an iterator.
    """
    # Search 1st occurence of a hex block:
    match = re_hexblock.search(data)
    if match is None:
        log.debug('No hex block found.')
        # no hex block found
        return
    while match is not None:
        found = match.group(0)
        # start index
        start = match.start()
        # current position
        current = match.end()
        log.debug('Found hex block starting at %08X, end %08X, size=%d' % (start, current, len(found)))
        if len(found) < min_size:
            log.debug('Too small - size<%d, ignored.' % min_size)
            match = re_hexblock.search(data, pos=current)
            continue
        #log.debug('Match: %s' % found)
        # remove all whitespace and line feeds:
        #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
        found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
        # TODO: make it a function
        # Also remove embedded RTF tags:
        found = re_embedded_tags.sub('', found)
        # object data extracted from the RTF file
        # MS Word accepts an extra hex digit, so we need to trim it if present:
        if len(found) & 1:
            log.debug('Odd length, trimmed last byte.')
            found = found[:-1]
        #log.debug('Cleaned match: %s' % found)
        objdata = binascii.unhexlify(found)
        # Detect the "\bin" control word, which is sometimes used for obfuscation:
        bin_match = re_delims_bin_decimal.match(data, pos=current)
        while bin_match is not None:
            log.debug('Found \\bin block starting at %08X : %r'
                          % (bin_match.start(), bin_match.group(0)))
            # extract the decimal integer following '\bin'
            bin_len = int(bin_match.group(1))
            log.debug('\\bin block length = %d' % bin_len)
            if current+bin_len > len(data):
                log.error('\\bin block length is larger than the remaining data')
                # move the current index, ignore the \bin block
                current += len(bin_match.group(0))
                break
            # read that number of bytes:
            objdata += data[current:current+bin_len]
            # TODO: handle exception
            current += len(bin_match.group(0)) + bin_len
            # TODO: check if current is out of range
            # TODO: is Word limiting the \bin length to a number of digits?
            log.debug('Current position = %08X' % current)
            match = re_delim_hexblock.match(data, pos=current)
            if match is not None:
                log.debug('Found next hex block starting at %08X, end %08X'
                    % (match.start(), match.end()))
                found = match.group(0)
                log.debug('Match: %s' % found)
                # remove all whitespace and line feeds:
                #NOTE: with Python 2.6+, we could use None instead of TRANSTABLE_NOCHANGE
                found = found.translate(TRANSTABLE_NOCHANGE, ' \t\r\n\f\v')
                # Also remove embedded RTF tags:
                found = re_embedded_tags.sub(found, '')
                objdata += binascii.unhexlify(found)
                current = match.end()
            bin_match = re_delims_bin_decimal.match(data, pos=current)

        # print repr(found)
        if len(objdata)>min_size:
            yield start, current-start, objdata
        # Search next occurence of a hex block:
        match = re_hexblock.search(data, pos=current)



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
    print '-'*79
    print 'File: %r - %d bytes' % (filename, len(data))
    for index, orig_len, objdata in rtf_iter_objects(data):
        print 'found object size %d at index %08X - end %08X' % (len(objdata), index, index+orig_len)
        fname = '%s_object_%08X.raw' % (fname_prefix, index)
        print 'saving object to file %s' % fname
        open(fname, 'wb').write(objdata)
        # TODO: check if all hex data is extracted properly

        obj = OleObject()
        try:
            obj.parse(objdata)
            print 'extract file embedded in OLE object:'
            print 'format_id  = %d' % obj.format_id
            print 'class name = %r' % obj.class_name
            print 'data size  = %d' % obj.data_size
            # set a file extension according to the class name:
            class_name = obj.class_name.lower()
            if class_name.startswith('word'):
                ext = 'doc'
            elif class_name.startswith('package'):
                ext = 'package'
            else:
                ext = 'bin'

            fname = '%s_object_%08X.%s' % (fname_prefix, index, ext)
            print 'saving to file %s' % fname
            open(fname, 'wb').write(obj.data)
            if obj.class_name.lower() == 'package':
                print 'Parsing OLE Package'
                opkg = OleNativeStream(bindata=obj.data)
                print 'Filename = %r' % opkg.filename
                print 'Source path = %r' % opkg.src_path
                print 'Temp path = %r' % opkg.temp_path
                if opkg.filename:
                    fname = '%s_%s' % (fname_prefix,
                                       sanitize_filename(opkg.filename))
                else:
                    fname = '%s_object_%08X.noname' % (fname_prefix, index)
                print 'saving to file %s' % fname
                open(fname, 'wb').write(opkg.data)
        except:
            pass
            log.exception('*** Not an OLE 1.0 Object')



#=== MAIN =================================================================

if __name__ == '__main__':
    # print banner with version
    print ('rtfobj %s - http://decalage.info/python/oletools' % __version__)
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
        print __doc__
        parser.print_help()
        sys.exit()

    # Setup logging to the console:
    # here we use stdout instead of stderr by default, so that the output
    # can be redirected properly.
    logging.basicConfig(level=LOG_LEVELS[options.loglevel], stream=sys.stdout,
                        format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)
    oleobj.log.setLevel(logging.NOTSET)


    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
        zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        process_file(container, filename, data, options.output_dir)




