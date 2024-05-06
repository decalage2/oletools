#!/usr/bin/env python3

""" Common operations for OpenXML files (docx, xlsx, pptx, ...)

This is mostly based on ECMA-376 (5th edition, Part 1)
http://www.ecma-international.org/publications/standards/Ecma-376.htm

See also: Notes on Microsoft's implementation of ECMA-376: [MS-0E376]

.. codeauthor:: Intra2net AG <info@intra2net>
License: BSD, see source code or documentation

ooxml is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE =================================================================

# ooxml is copyright (c) 2017-2020 Philippe Lagadec (http://www.decalage.info)
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

# -----------------------------------------------------------------------------
# CHANGELOG:
# 2018-12-06       CH: - ensure stdout can handle unicode

__version__ = '0.54.2'

# -- TODO ---------------------------------------------------------------------

# TODO: may have to tell apart single xml types: office2003 looks much different
#       than 2006+ --> DOCTYPE_*_XML2003
# TODO: check what is duplicate here with oleid, maybe merge some day?
# TODO: "xml2003" == "flatopc"? (No)


# -- IMPORTS ------------------------------------------------------------------

import sys
from oletools.common.log_helper import log_helper
from oletools.common.io_encoding import uopen
from zipfile import ZipFile, BadZipfile, is_zipfile
from os.path import splitext
import io
import re

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    import xml.etree.cElementTree as ET

###############################################################################
# CONSTANTS
###############################################################################


logger = log_helper.get_or_create_silent_logger('ooxml')

#: subfiles that have to be part of every ooxml file
FILE_CONTENT_TYPES = '[Content_Types].xml'
FILE_RELATIONSHIPS = '_rels/.rels'

#: start of content type attributes
CONTENT_TYPES_EXCEL = (
    'application/vnd.openxmlformats-officedocument.spreadsheetml.',
    'application/vnd.ms-excel.',
)
CONTENT_TYPES_WORD = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.',
)
CONTENT_TYPES_PPT = (
    'application/vnd.openxmlformats-officedocument.presentationml.',
)

#: other content types (currently unused)
CONTENT_TYPES_NEUTRAL = (
    'application/xml',
    'application/vnd.openxmlformats-package.relationships+xml',
    'application/vnd.openxmlformats-package.core-properties+xml',
    'application/vnd.openxmlformats-officedocument.theme+xml',
    'application/vnd.openxmlformats-officedocument.extended-properties+xml'
)

#: constants used to determine type of single-xml files
OFFICE_XML_PROGID_REGEX = r'<\?mso-application progid="(.*)"\?>'
WORD_XML_PROG_ID = 'Word.Document'
EXCEL_XML_PROG_ID = 'Excel.Sheet'

#: constants for document type
DOCTYPE_WORD = 'word'
DOCTYPE_EXCEL = 'excel'
DOCTYPE_POWERPOINT = 'powerpoint'
DOCTYPE_NONE = 'none'
DOCTYPE_MIXED = 'mixed'
DOCTYPE_WORD_XML = 'word-xml'
DOCTYPE_EXCEL_XML = 'excel-xml'
DOCTYPE_WORD_XML2003 = 'word-xml2003'
DOCTYPE_EXCEL_XML2003 = 'excel-xml2003'


###############################################################################
# HELPERS
###############################################################################


def debug_str(elem):
    """ for debugging: print an element """
    if elem is None:
        return u'None'
    if elem.tag[0] == '{' and elem.tag.count('}') == 1:
        parts = ['[tag={{...}}{0}'.format(elem.tag[elem.tag.index('}')+1:]), ]
    else:
        parts = ['[tag={0}'.format(elem.tag), ]
    if elem.text:
        parts.append(u'text="{0}"'.format(elem.text.replace('\n', '\\n')))
    if elem.tail:
        parts.append(u'tail="{0}"'.format(elem.tail.replace('\n', '\\n')))
    for key, value in elem.attrib.items():
        parts.append(u'{0}="{1}"'.format(key, value))
        if key == 'ContentType':
            if value.startswith(CONTENT_TYPES_EXCEL):
                parts[-1] += u'-->xls'
            elif value.startswith(CONTENT_TYPES_WORD):
                parts[-1] += u'-->doc'
            elif value.startswith(CONTENT_TYPES_PPT):
                parts[-1] += u'-->ppt'
            elif value in CONTENT_TYPES_NEUTRAL:
                parts[-1] += u'-->_'
            else:
                parts[-1] += u'!!!'

    text = u', '.join(parts)
    if len(text) > 150:
        return text[:147] + u'...]'
    return text + u']'


def isstr(some_var):
    """ version-independent test for isinstance(some_var, (str, unicode)) """
    if sys.version_info.major == 2:
        return isinstance(some_var, basestring)  # true for str and unicode   # pylint: disable=undefined-variable
    return isinstance(some_var, str)         # there is no unicode


###############################################################################
# INFO ON FILES
###############################################################################


def get_type(filename):
    """ return one of the DOCTYPE_* constants or raise error """
    parser = XmlParser(filename)
    if parser.is_single_xml():
        match = None
        with uopen(filename, 'r') as handle:
            match = re.search(OFFICE_XML_PROGID_REGEX, handle.read(1024))
        if not match:
            return DOCTYPE_NONE
        prog_id = match.groups()[0]
        if prog_id == WORD_XML_PROG_ID:
            return DOCTYPE_WORD_XML
        if prog_id == EXCEL_XML_PROG_ID:
            return DOCTYPE_EXCEL_XML
        return DOCTYPE_NONE

    is_doc = False
    is_xls = False
    is_ppt = False
    try:
        for _, elem, _ in parser.iter_xml(FILE_CONTENT_TYPES):
            logger.debug(u'  ' + debug_str(elem))
            try:
                content_type = elem.attrib['ContentType']
            except KeyError:         # ContentType not an attr
                continue
            is_xls |= content_type.startswith(CONTENT_TYPES_EXCEL)
            is_doc |= content_type.startswith(CONTENT_TYPES_WORD)
            is_ppt |= content_type.startswith(CONTENT_TYPES_PPT)
    except BadOOXML as oo_err:
        if oo_err.more_info.startswith('invalid subfile') and \
                FILE_CONTENT_TYPES in oo_err.more_info:
            # no FILE_CONTENT_TYPES in zip, so probably no ms office xml.
            return DOCTYPE_NONE
        raise

    if is_doc and not is_xls and not is_ppt:
        return DOCTYPE_WORD
    if not is_doc and is_xls and not is_ppt:
        return DOCTYPE_EXCEL
    if not is_doc and not is_xls and is_ppt:
        return DOCTYPE_POWERPOINT
    if not is_doc and not is_xls and not is_ppt:
        return DOCTYPE_NONE
    logger.warning('Encountered contradictory content types')
    return DOCTYPE_MIXED


def is_ooxml(filename):
    """ Determine whether given file is an ooxml file; tries get_type """
    try:
        doctype = get_type(filename)
    except BadOOXML:
        return False
    except IOError:   # one of the required files is not present
        return False
    if doctype == DOCTYPE_NONE:
        return False
    return True


###############################################################################
# HELPER CLASSES
###############################################################################


class ZipSubFile(object):
    """ A file-like object like ZipFile.open returns them, with size and seek()

    ZipFile.open() gives file handles that can be read but not seek()ed since
    the file is being decompressed in the background. This class implements a
    reset() function (close and re-open stream) and a seek() that uses it.
    --> can be used as argument to olefile.OleFileIO and olefile.isOleFile()

    Can be used as a context manager::

        with zipfile.ZipFile('file.zip') as zipper:
            # replaces with zipper.open(subfile) as handle:
            with ZipSubFile(zipper, 'subfile') as handle:
                print('subfile in file.zip has size {0}, starts with {1}'
                      .format(handle.size, handle.read(20)))
                handle.reset()

    Attributes always present:
    container: the containing zip file
    name: name of file within zip file
    mode: open-mode, 'r' per default
    size: size of the stream (constructor arg or taken from ZipFile.getinfo)
    closed: True if there was an open() but no close() since then

    Attributes only not-None after open() and before close():
    handle: direct handle to subfile stream, created by ZipFile.open()
    pos: current position within stream (can deviate from actual position in
         self.handle if we fake jump to end)

    See also (and maybe could some day merge with):
    ppt_record_parser.IterStream; also: oleobj.FakeFile
    """
    CHUNK_SIZE = 4096

    def __init__(self, container, filename, mode='r', size=None):
        """ remember all necessary vars but do not open yet """
        self.container = container
        self.name = filename
        if size is None:
            self.size = container.getinfo(filename).file_size
            logger.debug('zip stream has size {0}'.format(self.size))
        else:
            self.size = size
        if 'w' in mode.lower():
            raise ValueError('Can only read, mode "{0}" not allowed'
                             .format(mode))
        self.mode = mode
        self.handle = None
        self.pos = None
        self.closed = True

    def readable(self):
        return True

    def writable(self):
        return False

    def seekable(self):
        return True

    def open(self):
        """ open subfile for reading; open mode given to constructor before """
        if self.handle is not None:
            raise IOError('re-opening file not supported!')
        self.handle = self.container.open(self.name, self.mode)
        self.pos = 0
        self.closed = False
        # print('ZipSubFile: opened; size={}'.format(self.size))
        return self

    def write(self, *args, **kwargs):
        """ write is not allowed """
        raise IOError('writing not implemented')

    def read(self, size=-1):
        """
        read given number of bytes (or all data) from stream

        returns bytes (i.e. str in python2, bytes in python3)
        """
        if self.handle is None:
            raise IOError('read on closed handle')
        if self.pos >= self.size:
            # print('ZipSubFile: read fake at end')
            return b''   # fake being at the end, even if we are not
        data = self.handle.read(size)
        self.pos += len(data)
        # print('ZipSubFile: read {} bytes, pos now {}'.format(size, self.pos))
        return data

    def seek(self, pos, offset=io.SEEK_SET):
        """ re-position point so read() will continue elsewhere """
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

        # now get to that position, doing reads and resets as necessary
        if new_pos < 0:
            # print('ZipSubFile: Error: seek to {}'.format(new_pos))
            raise IOError('Seek beyond start of file not allowed')
        elif new_pos == self.pos:
            # print('ZipSubFile: nothing to do')
            pass
        elif new_pos == 0:
            # print('ZipSubFile: seek to start')
            self.reset()
        elif new_pos < self.pos:
            # print('ZipSubFile: seek back')
            self.reset()
            self._seek_skip(new_pos)             # --> read --> update self.pos
        elif new_pos < self.size:
            # print('ZipSubFile: seek forward')
            self._seek_skip(new_pos - self.pos)  # --> read --> update self.pos
        else:   # new_pos >= self.size
            # print('ZipSubFile: seek to end')
            self.pos = new_pos    # fake being at the end; remember pos >= size

    def _seek_skip(self, to_skip):
        """ helper for seek: skip forward by given amount using read() """
        # print('ZipSubFile: seek by skipping {} bytes starting at {}'
        #       .format(self.pos, to_skip))
        n_chunks, leftover = divmod(to_skip, self.CHUNK_SIZE)
        for _ in range(n_chunks):
            self.read(self.CHUNK_SIZE)    # just read and discard
        self.read(leftover)
        # print('ZipSubFile: seek by skipping done, pos now {}'
        #       .format(self.pos))

    def tell(self):
        """ inform about position of next read """
        # print('ZipSubFile: tell-ing we are at {}'.format(self.pos))
        return self.pos

    def reset(self):
        """ close and re-open """
        # print('ZipSubFile: resetting')
        self.close()
        self.open()

    def close(self):
        """ close file """
        # print('ZipSubFile: closing')
        if self.handle is not None:
            self.handle.close()
        self.pos = None
        self.handle = None
        self.closed = True

    def __enter__(self):
        """ start of context manager; opens the file """
        # print('ZipSubFile: entering context')
        self.open()
        return self

    def __exit__(self, *args, **kwargs):
        """ end of context manager; closes the file """
        # print('ZipSubFile: exiting context')
        self.close()

    def __str__(self):
        """ creates a nice textual representation for this object """
        if self.handle is None:
            status = 'closed'
        elif self.pos == 0:
            status = 'open, at start'
        elif self.pos >= self.size:
            status = 'open, at end'
        else:
            status = 'open, at pos {0}'.format(self.pos)

        return '[ZipSubFile {0} (size {1}, mode {2}, {3})]' \
               .format(self.name, self.size, self.mode, status)


class BadOOXML(ValueError):
    """ exception thrown if file is not an office XML file """

    def __init__(self, filename, more_info=None):
        """ create exception, remember filename and more_info """
        super(BadOOXML, self).__init__(
            '{0} is not an Office XML file{1}'
            .format(filename, ': ' + more_info if more_info else ''))
        self.filename = filename
        self.more_info = more_info


###############################################################################
# PARSING
###############################################################################


class XmlParser(object):
    """ parser for OOXML files

    handles two different types of files: "regular" OOXML files are zip
    archives that contain xml data and possibly other files in binary format.
    In Office 2003, Microsoft introduced another xml-based format, which uses
    a single xml file as data source. The content of these types is also
    different. Method :py:meth:`is_single_xml` tells them apart.
    """

    def __init__(self, filename):
        self.filename = filename
        self.did_iter_all = False
        self.subfiles_no_xml = set()
        self._is_single_xml = None

    def is_single_xml(self):
        """ determine whether this is "regular" ooxml or a single xml file

        Raises a BadOOXML if this is neither one or the other
        """
        if self._is_single_xml is not None:
            return self._is_single_xml

        if is_zipfile(self.filename):
            self._is_single_xml = False
            return False

        # find prog id in xml prolog
        match = None
        with uopen(self.filename, 'r') as handle:
            match = re.search(OFFICE_XML_PROGID_REGEX, handle.read(1024))
        if match:
            self._is_single_xml = True
            return True
        raise BadOOXML(self.filename, 'is no zip and has no prog_id')

    def iter_files(self, args=None):
        """
        Find files in zip or just give single xml file

        yields pairs (subfile-name, file-handle) where file-handle is an open
        file-like object. (Do not care too much about encoding here, the xml
        parser reads the encoding from the first lines in the file.)
        """
        if self.is_single_xml():
            if args:
                raise BadOOXML(self.filename, 'xml has no subfiles')
            # do not use uopen, xml parser determines encoding on its own
            with open(self.filename, 'rb') as handle:
                yield None, handle   # the subfile=None is needed in iter_xml
            self.did_iter_all = True
        else:
            zipper = None
            subfiles = None
            try:
                zipper = ZipFile(self.filename)
                if not args:
                    subfiles = zipper.namelist()
                elif isstr(args):
                    subfiles = [args, ]
                else:
                    # make a copy in case original args are modified
                    # Not sure whether this really is needed...
                    subfiles = tuple(arg for arg in args)

                for subfile in subfiles:
                    with zipper.open(subfile, 'r') as handle:
                        yield subfile, handle
                if not args:
                    self.did_iter_all = True
            except KeyError as orig_err:
                # Note: do not change text of this message without adjusting
                #       conditions in except handlers
                raise BadOOXML(self.filename,
                               'invalid subfile: ' + str(orig_err))
            except BadZipfile:
                raise BadOOXML(self.filename, 'not in zip format')
            finally:
                if zipper:
                    zipper.close()

    def iter_xml(self, subfiles=None, need_children=False, tags=None):
        """ Iterate xml contents of document

        If given subfile name[s] as optional arg[s], will only parse that
        subfile[s]

        yields 3-tuples (subfilename, element, depth) where depth indicates how
        deep in the hierarchy the element is located. Containers of element
        will come *after* the elements they contain (since they are only
        finished then).

        Subfiles that are not xml (e.g. OLE or image files) are remembered
        internally and can be retrieved using iter_non_xml().

        The argument need_children is set to False per default. If you need to
        access an element's children, set it to True. Note, however, that
        leaving it at False should save a lot of memory. Otherwise, the parser
        has to keep every single element in memory since the last element
        returned is the root which has the rest of the document as children.
        c.f. http://www.ibm.com/developerworks/xml/library/x-hiperfparse/

        Argument tags restricts output to tags with names from that list (or
        equal to that string). Children are preserved for these.
        """
        if tags is None:
            want_tags = []
        elif isstr(tags):
            want_tags = [tags, ]
            logger.debug('looking for tags: {0}'.format(tags))
        else:
            want_tags = tags
            logger.debug('looking for tags: {0}'.format(tags))

        for subfile, handle in self.iter_files(subfiles):
            events = ('start', 'end')
            depth = 0
            inside_tags = []
            try:
                for event, elem in ET.iterparse(handle, events):
                    if elem is None:
                        continue
                    if event == 'start':
                        if elem.tag in want_tags:
                            logger.debug('remember start of tag {0} at {1}'
                                         .format(elem.tag, depth))
                            inside_tags.append((elem.tag, depth))
                        depth += 1
                        continue
                    assert(event == 'end')
                    depth -= 1
                    assert(depth >= 0)

                    is_wanted = elem.tag in want_tags
                    if is_wanted:
                        curr_tag = (elem.tag, depth)
                        try:
                            if inside_tags[-1] == curr_tag:
                                inside_tags.pop()
                            else:
                                logger.error('found end for wanted tag {0} '
                                             'but last start tag {1} does not'
                                             ' match'.format(curr_tag,
                                                             inside_tags[-1]))
                                # try to recover: close all deeper tags
                                while inside_tags and \
                                        inside_tags[-1][1] >= depth:
                                    logger.debug('recover: pop {0}'
                                                 .format(inside_tags[-1]))
                                    inside_tags.pop()
                        except IndexError:    # no inside_tag[-1]
                            logger.error('found end of {0} at depth {1} but '
                                         'no start event')
                    # yield element
                    if is_wanted or not want_tags:
                        yield subfile, elem, depth

                    # save memory: clear elem so parser memorizes less
                    if not need_children and not inside_tags:
                        elem.clear()
                        # cannot do this since we might be using py-builtin xml
                        # while elem.getprevious() is not None:
                        #     del elem.getparent()[0]
            except ET.ParseError as err:
                self.subfiles_no_xml.add(subfile)
                if subfile is None:    # this is no zip subfile but single xml
                    raise BadOOXML(self.filename, 'content is not valid XML')
                elif subfile.endswith('.xml'):
                    log = logger.warning
                else:
                    log = logger.debug
                log('  xml-parsing for {0} failed ({1}). '
                    .format(subfile, err) +
                    'Run iter_non_xml to investigate.')
            assert(depth == 0)

    def get_content_types(self):
        """ retrieve subfile infos from [Content_Types].xml subfile

        returns (files, defaults) where
        - files is a dict that maps file-name --> content-type
        - defaults is a dict that maps extension --> content-type

        No guarantees on accuracy of these content types!
        """
        if self.is_single_xml():
            return {}, {}

        defaults = []
        files = []
        try:
            for _, elem, _ in self.iter_xml(FILE_CONTENT_TYPES):
                if elem.tag.endswith('Default'):
                    extension = elem.attrib['Extension']
                    if extension.startswith('.'):
                        extension = extension[1:]
                    defaults.append((extension, elem.attrib['ContentType']))
                    logger.debug('found content type for extension {0[0]}: '
                                 '{0[1]}'.format(defaults[-1]))
                elif elem.tag.endswith('Override'):
                    subfile = elem.attrib['PartName']
                    if subfile.startswith('/'):
                        subfile = subfile[1:]
                    files.append((subfile, elem.attrib['ContentType']))
                    logger.debug('found content type for subfile {0[0]}: '
                                 '{0[1]}'.format(files[-1]))
        except BadOOXML as oo_err:
            if oo_err.more_info.startswith('invalid subfile') and \
                    FILE_CONTENT_TYPES in oo_err.more_info:
                # no FILE_CONTENT_TYPES in zip, so probably no ms office xml.
                # Maybe OpenDocument format? In any case, try to analyze.
                pass
            else:
                raise
        return dict(files), dict(defaults)

    def iter_non_xml(self):
        """ retrieve subfiles that were found by iter_xml to be non-xml

        also looks for content type info in the [Content_Types].xml subfile.

        yields 3-tuples (filename, content_type, file_handle) where
        content_type is based on filename or default for extension or is None,
        and file_handle is a ZipSubFile. Caller does not have to care about
        closing handle, will be closed even in error condition.

        To handle binary parts of an xlsb file, use xls_parser.parse_xlsb_part
        """
        if not self.did_iter_all:
            logger.warning('Did not iterate through complete file. '
                           'Should run iter_xml() without args, first.')
        if not self.subfiles_no_xml:
            return

        # case of single xml files (office 2003+)
        if self.is_single_xml():
            return

        content_types, content_defaults = self.get_content_types()

        with ZipFile(self.filename) as zipper:
            for subfile in self.subfiles_no_xml:
                if subfile.startswith('/'):
                    subfile = subfile[1:]
                content_type = None
                if subfile in content_types:
                    content_type = content_types[subfile]
                else:
                    extension = splitext(subfile)[1]
                    if extension.startswith('.'):
                        extension = extension[1:]   # remove the '.'
                    if extension in content_defaults:
                        content_type = content_defaults[extension]
                with ZipSubFile(zipper, subfile) as handle:
                    yield subfile, content_type, handle


def test():
    """
    Test xml parsing; called when running this file as a script.

    Prints every element found in input file (to be given as command line arg).
    """
    log_helper.enable_logging(False, 'debug')
    if len(sys.argv) != 2:
        print(u'To test this code, give me a single file as arg')
        return 2

    # test get_type
    print('Detected type: ' + get_type(sys.argv[1]))

    # test complete parsing
    parser = XmlParser(sys.argv[1])
    for subfile, elem, depth in parser.iter_xml():
        if depth < 4:
            print(u'{0} {1}{2}'.format(subfile, '  ' * depth, debug_str(elem)))
    for index, (subfile, content_type, _) in enumerate(parser.iter_non_xml()):
        print(u'Non-XML subfile: {0} of type {1}'
              .format(subfile, content_type or u'unknown'))
        if index > 100:
            print(u'...')
            break

    log_helper.end_logging()

    return 0


if __name__ == '__main__':
    sys.exit(test())
