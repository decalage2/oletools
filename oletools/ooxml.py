#!/usr/bin/env python3

""" Common operations for OpenOffice XML (docx, xlsx, pptx, ...) files

This is mostly based on ECMA-376 (5th edition, Part 1)
http://www.ecma-international.org/publications/standards/Ecma-376.htm

See also: Notes on Microsoft's implementation of ECMA-376: [MS-0E376]

.. codeauthor:: Intra2net AG <info@intra2net>
"""

import sys
import logging
from zipfile import ZipFile, BadZipfile
from os.path import splitext
import io

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    import xml.etree.cElementTree as ET


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

#: constants for document type
DOCTYPE_WORD = 'word'
DOCTYPE_EXCEL = 'excel'
DOCTYPE_POWERPOINT = 'powerpoint'
DOCTYPE_NONE = 'none'
DOCTYPE_MIXED = 'mixed'


def debug_str(elem):
    """ for debugging: print an element """
    if elem is None:
        return u'None'
    if elem.tag[0] == '{' and elem.tag.count('}') == 1:
        parts = ['[tag={{...}}{0}'.format(elem.tag[elem.tag.index('}')+1:]), ]
    else:
        parts = ['[tag={0}'.format(elem.tag), ]
    if elem.text:
        parts.append(u'text="{0}"'.format(elem.text))
    if elem.tail:
        parts.append(u'tail="{0}"'.format(elem.tail))
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

    return u', '.join(parts) + u']'


def get_type(filename):
    """ return one of the DOCTYPE_* constants or raise error """
    is_doc = False
    is_xls = False
    is_ppt = False
    for _, elem, _ in XmlParser(filename).iter_xml(FILE_CONTENT_TYPES):
        logging.debug(u'  ' + debug_str(elem))
        try:
            is_xls |= elem.attrib['ContentType'].startswith(
                                            CONTENT_TYPES_EXCEL)
            is_doc |= elem.attrib['ContentType'].startswith(
                                            CONTENT_TYPES_WORD)
            is_ppt |= elem.attrib['ContentType'].startswith(
                                            CONTENT_TYPES_PPT)
        except KeyError:         # ContentType not an attr
            pass

    if is_doc and not is_xls and not is_ppt:
        return DOCTYPE_WORD
    if not is_doc and is_xls and not is_ppt:
        return DOCTYPE_EXCEL
    if not is_doc and not is_xls and is_ppt:
        return DOCTYPE_POWERPOINT
    if not is_doc and not is_xls and not is_ppt:
        return DOCTYPE_NONE
    else:
        return DOCTYPE_MIXED


def is_ooxml(filename):
    """ Determine whether given file is an ooxml file; tries get_type """
    try:
        get_type(filename)
    except BadZipfile:
        return False
    except IOError:   # one of the required files is not present
        return False


class ZipSubFile(object):
    """ A file-like object like ZipFile.open returns them, with size and seek()

    ZipFile.open() gives file handles that can be read but not seek()ed since
    the file is being decompressed in the background. This class implements a
    reset() function which corresponds to a seek to 0 (which just closes the
    stream and re-opens it behind the scenes.)
    --> can be used e.g. for olefile.isOleFile()

    Can be used as a context manager::

        with zipfile.ZipFile('file.zip') as zipper:
            with ZipSubFile(zipper, 'subfile') as handle:
                print('subfile in file.zip has size {0}, starts with {1}'
                      .format(handle.size, handle.read(20)))
                handle.reset()

    Attributes always present:
    container: the containing zip file
    name: name of file within zip file
    mode: open-mode, 'r' per default
    size: size of the stream (constructor arg or taken from ZipFile.getinfo)

    Attributes only not-None after open() and before close():
    handle: direkt handle to subfile stream, created by ZipFile.open()
    pos: current position within stream
    """

    def __init__(self, container, filename, mode='r', size=None):
        """ remember all necessary vars but do not open yet """
        self.container = container
        self.name = filename
        if size is None:
            self.size = container.getinfo(filename).file_size
            logging.debug('zip stream has size {0}'.format(self.size))
        else:
            self.size = size
        if 'w' in mode.lower():
            raise ValueError('Can only read, mode "{0}" not allowed'
                             .format(mode))
        self.mode = mode
        self.handle = None
        self.pos = None

    def open(self):
        """ open subfile for reading; open mode given to constructor before """
        if self.handle is not None:
            raise IOError('re-opening file not supported!')
        self.handle = self.container.open(self.name, self.mode)
        self.pos = 0
        return self

    def write(self, *args, **kwargs):                          # pylint: disable=unused-argument,no-self-use
        """ write is not allowed """
        raise IOError('writing not implemented')

    def read(self, size=-1):
        """
        read given number of bytes (or all data) from stream

        returns bytes (i.e. str in python2, bytes in python3)
        """
        if size is None:
            self.pos = self.size
        else:
            self.pos += size
        return self.handle.read(size)

    def seek(self, pos, offset):
        """ re-position point so read() will continue elsewhere

        only re-positioning to start of file is allowed
        """
        if pos == 0 and offset == io.SEEK_SET:
            self.reset()
        elif pos == -self.pos and offset == io.SEEK_CUR:
            self.reset()
        else:
            raise NotImplementedError('could reset() and read()')

    def tell(self):
        """ inform about position of next read """
        return self.pos

    def reset(self):
        """ close and re-open """
        self.close()
        self.open()

    def close(self):
        """ close file """
        if self.handle is not None:
            self.handle.close()
        self.pos = None
        self.handle = None

    def __enter__(self):
        """ start of context manager; opens the file """
        self.open()
        return self

    def __exit__(self, *args, **kwargs):
        """ end of context manager; closes the file """
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


class XmlParser(object):
    """ parser for OOXML files """

    def __init__(self, filename):
        self.filename = filename
        self.did_iter_all = False
        self.subfiles_no_xml = set()

    def iter_xml(self, *args):
        """ Iterate xml contents of document

        If given subfile name[s] as optional arg[s], will only parse that
        subfile[s]

        yields 3-tuples (subfilename, element, depth) where depth indicates how
        deep in the hierarchy the element is located. Containers of element
        will come *after* the elements they contain (since they are only
        finished then).

        Subfiles that are not xml (e.g. OLE or image files) are remembered
        internally and can be retrieved using iter_non_xml().
        """
        with ZipFile(self.filename) as zipper:
            if args:
                subfiles = args
            else:
                subfiles = zipper.namelist()

            events = ('start', 'end')
            for subfile in subfiles:
                logging.debug(u'subfile {0}'.format(subfile))
                depth = 0
                try:
                    with zipper.open(subfile, 'r') as handle:
                        for event, elem in ET.iterparse(handle, events):
                            if elem is None:
                                continue
                            if event == 'start':
                                depth += 1
                                continue
                            assert(event == 'end')
                            depth -= 1
                            assert(depth >= 0)
                            yield subfile, elem, depth
                except ET.ParseError as err:
                    if subfile.endswith('.xml'):
                        logger = logging.warning
                    else:
                        logger = logging.debug
                    logger('  xml-parsing for {0} failed ({1}). '
                           .format(subfile, err) +
                           'Run iter_non_xml to investigate.')
                    self.subfiles_no_xml.add(subfile)
                assert(depth == 0)
        if not args:
            self.did_iter_all = True

    def get_content_types(self):
        """ retrieve subfile infos from [Content_Types].xml subfile

        returns (files, defaults) where
        - files is a dict that maps file-name --> content-type
        - defaults is a dict that maps extension --> content-type

        No guarantees on accuracy of these content types!
        """
        defaults = []
        files = []
        for _, elem, _ in self.iter_xml(FILE_CONTENT_TYPES):
            if elem.tag.endswith('Default'):
                extension = elem.attrib['Extension']
                if extension.startswith('.'):
                    extension = extension[1:]
                defaults.append((extension, elem.attrib['ContentType']))
                logging.debug('found content type for extension {0[0]}: {0[1]}'
                              .format(defaults[-1]))
            elif elem.tag.endswith('Override'):
                subfile = elem.attrib['PartName']
                if subfile.startswith('/'):
                    subfile = subfile[1:]
                files.append((subfile, elem.attrib['ContentType']))
                logging.debug('found content type for subfile {0[0]}: {0[1]}'
                              .format(files[-1]))
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
            logging.warning('Did not iterate through complete file. '
                            'Should run iter_xml() without args, first.')
        if not self.subfiles_no_xml:
            raise StopIteration()

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
    """ Main function, called when running file as script

    see module doc for more info
    """
    logging.basicConfig(level=logging.DEBUG)
    if len(sys.argv) != 2:
        print(u'To test this code, give me a single file as arg')
        return 2
    parser = XmlParser(sys.argv[1])
    for subfile, elem, depth in parser.iter_xml():
        print(u'{0}{1}{2}'.format(subfile, '  ' * depth, debug_str(elem)))
    for subfile, content_type in parser.iter_non_xml():
        print(u'Non-XML subfile: {0} of type {1}'
              .format(subfile, content_type or u'unknown'))
    return 0


if __name__ == '__main__':
    sys.exit(test())
