#!/usr/bin/env python
"""
ftguess.py

ftguess is a Python module to determine the type of a file based on its contents.
It can be used as a Python library or a command-line tool.

Usage: ftguess <file>

ftguess is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# Useful resources about file formats:
# http://fileformats.archiveteam.org
# https://www.nationalarchives.gov.uk/PRONOM/Default.aspx

#=== LICENSE =================================================================

# ftguess is copyright (c) 2018-2024, Philippe Lagadec (http://www.decalage.info)
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

from __future__ import print_function

#------------------------------------------------------------------------------
# CHANGELOG:
# 2018-07-04 v0.54 PL: - first version
# 2021-05-09 v0.60 PL: -

__version__ = '0.60.2'

# ------------------------------------------------------------------------------
# TODO:


# === IMPORTS =================================================================

import sys
import io
import zipfile
import os
import olefile
import logging
import optparse

# import lxml or ElementTree for XML parsing:
try:
    # lxml: best performance for XML processing
    import lxml.etree as ET
except ImportError:
    try:
        # Python 2.5+: batteries included
        import xml.etree.ElementTree as ET
    except ImportError:
        try:
            # Python <2.5: standalone ElementTree install
            import elementtree.cElementTree as ET
        except ImportError:
            raise ImportError("lxml or ElementTree are not installed, " \
                               + "see http://codespeak.net/lxml " \
                               + "or http://effbot.org/zone/element-index.htm")

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

from oletools.common import clsid
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
log = get_logger('ftguess')

def enable_logging():
    """
    Enable logging for this module (disabled by default).
    This will set the module-specific logger level to NOTSET, which
    means the main application controls the actual logging level.
    """
    log.setLevel(logging.NOTSET)

# === CONSTANTS ===============================================================

# file types for FileTypeGuesser:
class FTYPE(object):
    """
    Constants for file types
    """
    ZIP = 'Zip'
    WORD = 'Word'
    WORD6 = 'Word6'
    WORD97 = 'Word97'
    WORD2007 = 'Word2007'
    WORD2007_DOCX = 'Word2007_DOCX'
    WORD2007_DOTX = 'Word2007_DOTX'
    WORD2007_DOCM = 'Word2007_DOCM'
    WORD2007_DOTM = 'Word2007_DOTM'
    EXCEL = 'Excel'
    EXCEL5 = 'Excel5'
    EXCEL97 = 'Excel97'
    EXCEL2007 = 'Excel2007'
    EXCEL2007_XLSX = 'Excel2007_XLSX'
    EXCEL2007_XLSM = 'Excel2007_XLSM'
    EXCEL2007_XLTX = 'Excel2007_XLTX'
    EXCEL2007_XLTM = 'Excel2007_XLTM'
    EXCEL2007_XLSB = 'Excel2007_XLSB'
    EXCEL2007_XLAM = 'Excel2007_XLAM'
    POWERPOINT97 = 'Powerpoint97'
    POWERPOINT2007 = 'Powerpoint2007'
    POWERPOINT2007_PPTX = 'Powerpoint2007_PPTX'
    POWERPOINT2007_PPSX = 'Powerpoint2007_PPSX'
    POWERPOINT2007_PPTM = 'Powerpoint2007_PPTM'
    POWERPOINT2007_PPSM = 'Powerpoint2007_PPSM'
    # TODO: DOCM, PPTM, PPSX, PPSM, ...
    XPS = 'XPS'
    RTF = 'RTF'
    HTML = 'HTML'
    PDF = 'PDF'
    MHTML = 'MHTML'
    TEXT = 'TEXT'
    EXE_PE = 'EXE_PE'
    GENERIC_OLE = 'OLE' # Generic OLE file
    GENERIC_XML = 'XML' # Generic XML file
    GENERIC_OPENXML = 'OpenXML' # Generic OpenXML file
    UNKNOWN = 'Unknown File Type'
    MSI = "MSI"
    ONENOTE = "OneNote"
    PNG = 'PNG'

class CONTAINER(object):
    """
    Constants for file container types
    """
    RTF = 'RTF'
    ZIP = 'Zip'
    OLE = 'OLE'
    OpenXML = 'OpenXML'
    FlatOPC = 'FlatOPC'
    OpenDocument = 'OpenDocument'
    MIME = 'MIME'
    BINARY = 'Binary'  # Generic binary file without container
    UNKNOWN = 'Unknown Container'
    ONENOTE = 'OneNote'
    PNG = 'PNG'

class APP(object):
    """
    Constants for file types
    """
    MSWORD = 'MS Word'
    MSEXCEL = 'MS Excel'
    MSPOWERPOINT = 'MS PowerPoint'
    MSACCESS = 'MS Access'
    MSVISIO = 'MS Visio'
    MSPROJECT = 'MS Project'
    MSOFFICE = 'MS Office'  # when the exact app is unknown
    MSONENOTE = 'MS OneNote'
    ZIP_ARCHIVER = 'Any Zip Archiver'
    WINDOWS = 'Windows'  # for Windows executables and XPS
    UNKNOWN = 'Unknown Application'

# FTYPE_NAME = {
#     FTYPE_ZIP: 'Zip archive',
#     FTYPE_WORD97: 'MS Word 97-2000 Document',
# }

# Namespaces and tags for OpenXML parsing`- RELS files:
# root: <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
NS_RELS = '{http://schemas.openxmlformats.org/package/2006/relationships}'
TAG_RELS = NS_RELS + 'Relationships'
# <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/>
TAG_REL = NS_RELS + 'Relationship'
ATTR_REL_TYPE = 'Type'
ATTR_REL_TARGET = 'Target'
URL_REL_OFFICEDOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
# For "strict" OpenXML formats, the URL is different:
URL_REL_OFFICEDOC_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
# Url for xps files
URL_REL_XPS = 'http://schemas.microsoft.com/xps/2005/06/fixedrepresentation'
# Namespaces and tags for OpenXML parsing`- Content-types file:
NS_CONTENT_TYPES = '{http://schemas.openxmlformats.org/package/2006/content-types}'
TAG_CTYPES_DEFAULT = NS_CONTENT_TYPES + 'Default'
TAG_CTYPES_OVERRIDE = NS_CONTENT_TYPES + 'Override'


# Namespaces and tags for Word/PowerPoint 2007+ XML parsing:
# root: <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
NS_XMLPACKAGE = '{http://schemas.microsoft.com/office/2006/xmlPackage}'
TAG_PACKAGE = NS_XMLPACKAGE + 'package'
# the tag <pkg:part> includes <pkg:binaryData> that contains the VBA macro code in Base64:
# <pkg:part pkg:name="/word/vbaProject.bin" pkg:contentType="application/vnd.ms-office.vbaProject"><pkg:binaryData>
TAG_PKGPART = NS_XMLPACKAGE + 'part'
ATTR_PKG_NAME = NS_XMLPACKAGE + 'name'
ATTR_PKG_CONTENTTYPE = NS_XMLPACKAGE + 'contentType'
CTYPE_VBAPROJECT = "application/vnd.ms-office.vbaProject"
TAG_PKGBINDATA = NS_XMLPACKAGE + 'binaryData'




# === CLASSES ================================================================

class FType_Base (object):
    container = CONTAINER.UNKNOWN
    application = APP.UNKNOWN
    filetype = FTYPE.UNKNOWN
    name = "Unknown file type"
    longname = "Unknown file type"
    extensions = []  # list of common file extensions used for the format
    content_types = []  # list of MIME content-types (can be several)
    PUID = None  # PRONOM Unique ID - see https://www.nationalarchives.gov.uk/PRONOM/Default.aspx
    may_contain_vba = False
    may_contain_xlm = False
    may_contain_ole = False

    @classmethod
    def recognize(cls, ftg):
        """
        return True if the provided file matches the type of this class
        :param ftg: FileTypeGuesser object
        :return: bool
        """
        return False

class FType_Unknown(FType_Base):
    pass

class FType_RTF(FType_Base):
    container = CONTAINER.RTF
    application = APP.MSWORD
    filetype = FTYPE.RTF
    name = 'RTF'
    longname = 'Rich Text Format'
    extensions = ['rtf', 'doc']
    content_types = ('application/rtf', 'text/rtf')
    PUID = 'fmt/355'  # RTF 1.9 (from Word 2007)

    @classmethod
    def recognize(cls, ftg):
        # print('checking RTF')
        # print(repr(data[0:4]))
        return True if ftg.data.startswith(b'{\\rt') else False


class FType_Generic_OLE(FType_Base):
    container = CONTAINER.OLE
    application = APP.UNKNOWN
    filetype = FTYPE.GENERIC_OLE
    name = 'Generic OLE/CFB file'
    longname = 'Generic OLE file / Compound File (unknown format)'

    @classmethod
    def recognize(cls, ftg):
        # Here there's an issue with non-OLE files smaller than 1536 bytes
        # see https://github.com/decalage2/olefile/issues/142
        # Workaround: pad data when it's smaller than 1536 bytes
        # TODO: use the new data parameter of isOleFile when it's implemented
        if len(ftg.data)<1536:
            data = ftg.data + (b'\x00'*1536)
        else:
            data = ftg.data
        if olefile.isOleFile(data):
            # open the OLE file
            try:
                # Open and parse the OLE file:
                ftg.olefile = olefile.OleFileIO(ftg.data)
                # Extract the CLSID of the root storage
                ftg.root_clsid = ftg.olefile.root.clsid
                ftg.root_clsid_name = clsid.KNOWN_CLSIDS.get(ftg.root_clsid, None)
            except:
                # TODO: log the error
                return False
            return True
        else:
            return False


class FType_OLE_CLSID_Base(FType_Generic_OLE):
    """
    Base class to recognize OLE files based on CLSID or stream names
    """
    CLSIDS = []
    STREAMS = []

    @classmethod
    def recognize(cls, ftg):
        # TODO: refactor, this is not used anymore
        if ftg.root_clsid is not None:
            # First, attempt to identify the root storage CLSID:
            if ftg.root_clsid in cls.CLSIDS:
                return True
            else:
                return False
        else:
            # Second, check the presence of well-known stream names
            # TODO: check if a Word doc is OK without a clsid
            return False

class FType_Generic_Zip(FType_Base):
    container = CONTAINER.ZIP
    application = APP.ZIP_ARCHIVER
    filetype = FTYPE.ZIP
    name = 'Zip Archive'
    longname = 'Generic Zip Archive'
    extensions = ['zip']

    @classmethod
    def recognize(cls, ftg):
        # First, call is_zipfile to discard non zip archives:
        if not zipfile.is_zipfile(ftg.data_bytesio):
            return False
        # Second, attempt to open the zip file for further processing:
        try:
            ftg.zipfile = zipfile.ZipFile(ftg.data_bytesio)
        except zipfile.BadZipfile:
            # this exception happens only when the zip file could not be opened properly
            # it should not catch other potential errors
            return False
        return True


class FType_Generic_OpenXML(FType_Base):
    container = CONTAINER.OpenXML
    application = APP.MSOFFICE
    filetype = FTYPE.GENERIC_OPENXML
    name = 'OpenXML file'
    longname = 'Generic OpenXML file'
    extensions = []

    @classmethod
    def recognize(cls, ftg):
        log.debug('Open XML - recognize')
        # TODO: move most of this code to ooxml.py
        # TODO: here it can be either forward or backward slash...
        try:
            ftg.zipfile.getinfo('_rels/.rels')
        except KeyError:
            return False
        try:
            root_rels = ftg.zipfile.read('_rels/.rels')
        except RuntimeError:
            return False
        # parse the XML content
        # TODO: handle XML parsing exceptions
        elem_rels = ET.fromstring(root_rels)
        # check root:
        if elem_rels.tag != TAG_RELS:
            return False
        main_part = None
        for elem_rel in elem_rels.iter(tag=TAG_REL):
            rel_type = elem_rel.get(ATTR_REL_TYPE)
            log.debug('Relationship: type=%s target=%s' % (rel_type, elem_rel.get(ATTR_REL_TARGET)))
            if rel_type in (URL_REL_OFFICEDOC, URL_REL_OFFICEDOC_STRICT, URL_REL_XPS):
                # TODO: is it useful to distinguish normal and strict OpenXML?
                main_part = elem_rel.get(ATTR_REL_TARGET)
                # TODO: raise anomaly if there are more than one rel with type office doc
                break
        log.debug('Main part: %s' % main_part)
        # if main_part is not None:
        #     try:
        #         main_part_xml = ftg.zipfile.read(main_part)
        #     except RuntimeError:
        #         return False
        #     # TODO: handle XML parsing exceptions
        #     elem_main_part = ET.fromstring(main_part_xml)
        #     #print(elem_main_part.tag)
        #     # Save XML tag of main part to determine actual format
        #     ftg.main_part_xmltag = elem_main_part.tag
        # else:
        #     # TODO: log error, raise anomaly (or maybe it's the case for XPS?)
        #     return False
        if main_part is None:
            # just warn but do not raise an exception. This might be just
            # another strange data type out there that we do not understand
            # yet. Return False so file type will stay FType_Generic_OpenXML
            log.warning('Failed to find any known relationship in OpenXML-file')
            # TODO: here we should recognize a generic OpenXML type instead of returning False
            return False

        # parse content types, find content type of main part
        try:
            content_types = ftg.zipfile.read('[Content_Types].xml')
        except RuntimeError:
            return False
        # parse the XML content
        # TODO: handle XML parsing exceptions
        elem_ctypes = ET.fromstring(content_types)
        ctypes_ext = {}
        ctypes_part = {}
        for elem_ext in elem_ctypes.iter(tag = TAG_CTYPES_DEFAULT):
            extension = elem_ext.get('Extension')
            content_type = elem_ext.get('ContentType')
            # print('Ext: %s => Content-type: %s' % (extension, content_type))
            if extension is not None and content_type is not None:
                ctypes_ext[extension] = content_type
        for elem_part in elem_ctypes.iter(tag = TAG_CTYPES_OVERRIDE):
            partname = elem_part.get('PartName')
            # remove leading slash if present
            partname = partname.lstrip('/')
            content_type = elem_part.get('ContentType')
            # print('Part: %s => Content-type: %s' % (partname, content_type))
            if partname is not None and content_type is not None:
                ctypes_part[partname] = content_type
        # find content-type of the main part, first by part name, second by extension:
        main_part_content_type = None
        if main_part in ctypes_part:
            main_part_content_type = ctypes_part[main_part]
        else:
            # extract extension from part name, without leading dot
            main_part_ext = os.path.splitext(main_part)[1][1:]
            if main_part_ext in ctypes_ext:
                main_part_content_type = ctypes_ext[main_part_ext]
        ftg.main_part_content_type = main_part_content_type
        log.debug('Main part content-type: %s' % main_part_content_type)
        return True


# --- WORD Formats ---

class FType_Word(FType_Base):
    '''Base class for all MS Word file types'''
    application = APP.MSWORD
    name = 'MS Word (generic)'
    longname = 'MS Word Document or Template (generic)'

class FType_Word97(FType_OLE_CLSID_Base, FType_Word):
    application = APP.MSWORD
    filetype = FTYPE.WORD97
    name = 'MS Word 97 Document'
    longname = 'MS Word 97-2003 Document or Template'
    CLSIDS = ('00020906-0000-0000-C000-000000000046',)
    extensions = ['doc', 'dot']
    content_types = ['application/msword']
    PUID = 'fmt/40'
    may_contain_vba = True
    may_contain_ole = True
    # TODO: if no CLSID, check stream 'WordDocument'

class FType_Word6(FType_OLE_CLSID_Base, FType_Word):
    application = APP.MSWORD
    filetype = FTYPE.WORD6
    name = 'MS Word 6 Document'
    longname = 'MS Word 6-7 Document or Template'
    CLSIDS = ('00020900-0000-0000-C000-000000000046',)
    extensions = ['doc', 'dot']
    content_types = ['application/msword']
    PUID = 'fmt/39'
    may_contain_ole = True

class FType_Word2007_Base(FType_Generic_OpenXML, FType_Word):
    application = APP.MSWORD
    name = 'MS Word 2007+ File'
    longname = 'MS Word 2007+ File (.doc?)'


class FType_Word2007(FType_Word2007_Base):
    application = APP.MSWORD
    filetype = FTYPE.WORD2007_DOCX
    name = 'MS Word 2007+ Document'
    longname = 'MS Word 2007+ Document (.docx)'
    extensions = ['docx']

class FType_Word2007_Macro(FType_Word2007_Base):
    application = APP.MSWORD
    filetype = FTYPE.WORD2007_DOCM
    name = 'MS Word 2007+ Macro-Enabled Document'
    longname = 'MS Word 2007+ Macro-Enabled Document (.docm)'
    extensions = ['docm']

class FType_Word2007_Template(FType_Word2007_Base):
    application = APP.MSWORD
    filetype = FTYPE.WORD2007_DOTX
    name = 'MS Word 2007+ Template'
    longname = 'MS Word 2007+ Template (.dotx)'
    extensions = ['dotx']

class FType_Word2007_Template_Macro(FType_Word2007_Base):
    application = APP.MSWORD
    filetype = FTYPE.WORD2007_DOTM
    name = 'MS Word 2007+ Macro-Enabled Template'
    longname = 'MS Word 2007+ Macro-Enabled Template (.dotm)'
    extensions = ['dotm']

# --- EXCEL Formats ---

class FType_Excel(FType_Base):
    '''Base class for all MS Excel file types'''
    application = APP.MSEXCEL
    name = 'MS Excel (generic)'
    longname = 'MS Excel Workbook/Template/Add-in (generic)'

class FType_Excel97(FType_Excel, FType_Generic_OLE):
    filetype = FTYPE.EXCEL97
    name = 'MS Excel 97 Workbook'
    longname = 'MS Excel 97-2003 Workbook or Template'
    CLSIDS = ('00020820-0000-0000-C000-000000000046',)
    extensions = ['xls', 'xlt', 'xla']
    # TODO: if no CLSID, check stream 'Workbook' or 'Book' (maybe Excel 5)

class FType_Excel5(FType_Excel, FType_Generic_OLE):
    filetype = FTYPE.EXCEL5
    name = 'MS Excel 5.0/95 Workbook'
    longname = 'MS Excel 5.0/95 Workbook, Template or Add-in'
    CLSIDS = ('00020810-0000-0000-C000-000000000046',)
    extensions = ['xls', 'xlt', 'xla']
    # TODO: this CLSID is also used in Excel addins (.xla) saved by MS Excel 365

class FType_Excel2007(FType_Excel, FType_Generic_OpenXML):
    '''Base class for all MS Excel 2007 file types'''
    name = 'MS Excel 2007+ (generic)'
    longname = 'MS Excel 2007+ Workbook or Template (generic)'
    content_types = ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',)
      # note: content type differs only for xlsm

class FType_Excel2007_XLSX (FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLSX
    name = 'MS Excel 2007+ Workbook'
    longname = 'MS Excel 2007+ Workbook (.xlsx)'
    extensions = ['xlsx']
    PUID = 'fmt/214'

class FType_Excel2007_XLSM (FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLSM
    name = 'MS Excel 2007+ Macro-Enabled Workbook'
    longname = 'MS Excel 2007+ Macro-Enabled Workbook (.xlsm)'
    extensions = ['xlsm']
    content_types = ('application/vnd.ms-excel.sheet.macroEnabled.12',)
    PUID = 'fmt/445'

class FType_Excel2007_XLSB (FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLSB
    name = 'MS Excel 2007+ Binary Workbook'
    longname = 'MS Excel 2007+ Binary Workbook (.xlsb)'
    extensions = ['xlsb']
    content_types = ('application/vnd.ms-excel.sheet.binary.macroEnabled.12',)
    PUID = 'fmt/595'

class FType_Excel2007_Template(FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLTX
    name = 'MS Excel 2007+ Template'
    longname = 'MS Excel 2007+ Template (.xltx)'
    extensions = ['xltx']

class FType_Excel2007_Template_Macro(FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLTM
    name = 'MS Excel 2007+ Macro-Enabled Template'
    longname = 'MS Excel 2007+ Macro-Enabled Template (.xltm)'
    extensions = ['xltm']

class FType_Excel2007_Addin_Macro(FType_Excel2007):
    filetype = FTYPE.EXCEL2007_XLAM
    name = 'MS Excel 2007+ Macro-Enabled Add-in'
    longname = 'MS Excel 2007+ Macro-Enabled Add-in (.xlam)'
    extensions = ['xlam']

# --- POWERPOINT Formats ---

class FType_Powerpoint(FType_Base):
    '''Base class for all MS Powerpoint file types'''
    application = APP.MSPOWERPOINT
    name = 'MS Powerpoint (generic)'
    longname = 'MS Powerpoint Presentation/Slideshow/Template/Addin/... (generic)'

class FType_Powerpoint97(FType_Powerpoint, FType_Generic_OLE):
    # see also: ppt_record_parser.is_ppt
    filetype = FTYPE.POWERPOINT97
    name = 'MS Powerpoint 97 Presentation'
    longname = 'MS Powerpoint 97-2003 Presentation/Slideshow/Template'
    CLSIDS = ('64818D10-4F9B-11CF-86EA-00AA00B929E8',)
    extensions = ['ppt', 'pps', 'pot']

class FType_Powerpoint2007(FType_Powerpoint, FType_Generic_OpenXML):
    '''Base class for all MS Powerpoint 2007 file types'''
    filetype = FTYPE.POWERPOINT2007
    name = 'MS Powerpoint 2007+ (generic)'
    longname = 'MS Powerpoint 2007+ Presentation/Slideshow/Template (generic)'
    content_types = ('application/vnd.openxmlformats-officedocument.presentationml.presentation',)

class FType_Powerpoint2007_Presentation(FType_Powerpoint2007):
    filetype = FTYPE.POWERPOINT2007_PPTX
    name = 'MSPowerpoint 2007+ Presentation'
    longname = 'MSPowerpoint 2007+ Presentation (.pptx)'
    content_types = ('application/vnd.openxmlformats-officedocument.presentationml.presentation',)
    extensions = ['pptx']

class FType_Powerpoint2007_Slideshow(FType_Powerpoint2007):
    filetype = FTYPE.POWERPOINT2007_PPSX
    name = 'MSPowerpoint 2007+ Slideshow'
    longname = 'MSPowerpoint 2007+ Slideshow (.ppsx)'
    content_types = ('application/vnd.openxmlformats-officedocument.presentationml.slideshow',)
    extensions = ['ppsx']

class FType_Powerpoint2007_Macro(FType_Powerpoint2007):
    filetype = FTYPE.POWERPOINT2007_PPTM
    name = 'MSPowerpoint 2007+ Macro-Enabled Presentation'
    longname = 'MSPowerpoint 2007+ Macro-Enabled Presentation (.pptm)'
    content_types = ('application/vnd.ms-powerpoint.presentation.macroEnabled.12',)
    extensions = ['pptm']

class FType_Powerpoint2007_Slideshow_Macro(FType_Powerpoint2007):
    filetype = FTYPE.POWERPOINT2007_PPSM
    name = 'MSPowerpoint 2007+ Macro-Enabled Slideshow'
    longname = 'MSPowerpoint 2007+ Macro-Enabled Slideshow (.ppsm)'
    content_types = ('application/vnd.ms-powerpoint.slideshow.macroEnabled.12',)
    extensions = ['ppsm']


class FType_XPS(FType_Generic_OpenXML):
    application = APP.WINDOWS
    filetype = FTYPE.XPS
    name = 'XPS'
    longname = 'Fixed-Page Document (.xps)',
    extensions = ['xps']


class FType_MSI(FType_Generic_OLE):
    # see http://fileformats.archiveteam.org/wiki/Windows_Installer
    application = APP.WINDOWS
    filetype = FTYPE.MSI
    name = 'MSI'
    longname = 'Windows Installer Package (.msi)'
    extensions = ['msi']


class FType_OneNote(FType_Base):
    container = CONTAINER.ONENOTE
    application = APP.MSONENOTE
    filetype = FTYPE.ONENOTE
    name = 'OneNote'
    longname = 'MS OneNote Revision Store (.one)'
    extensions = ['one']
    content_types = ('application/msonenote',)
    PUID = 'fmt/637'
    # ref: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-onestore/ae670cd2-4b38-4b24-82d1-87cfb2cc3725
    # PRONOM: https://www.nationalarchives.gov.uk/PRONOM/Format/proFormatSearch.aspx?status=detailReport&id=1437

    @classmethod
    def recognize(cls, ftg):
        # ref about Header with OneNote GUID:
        # https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-onestore/2b394c6b-8788-441f-b631-da1583d772fd
        return True if ftg.data.startswith(b'\xE4\x52\x5C\x7B\x8C\xD8\xA7\x4D\xAE\xB1\x53\x78\xD0\x29\x96\xD3') else False


class FType_PNG(FType_Base):
    container = CONTAINER.PNG
    application = APP.UNKNOWN
    filetype = FTYPE.PNG
    name = 'PNG'
    longname = 'Portable Network Graphics picture (.png)'
    extensions = ['png']
    content_types = ('image/png',)
    PUID = 'fmt/13' # This is for PNG 1.2. PNG 1.1 is fmt/12, 1.0 is fmt/11
    # ref: http://fileformats.archiveteam.org/wiki/PNG
    # PRONOM: https://www.nationalarchives.gov.uk/PRONOM/Format/proFormatSearch.aspx?status=detailReport&id=666

    @classmethod
    def recognize(cls, ftg):
        return True if ftg.data.startswith(b'\x89\x50\x4E\x47\x0D\x0A\x1A\x0A') else False


# TODO: for PPT, check for stream 'PowerPoint Document'
# TODO: for Visio, check for stream 'VisioDocument'

clsid_ftypes = {
    # mapping from CLSID of root storage to FType classes:
    # TODO: do not repeat magic numbers, import from oletools.common.clsid
    # WORD
    '00020906-0000-0000-C000-000000000046': FType_Word97,
    '00020900-0000-0000-C000-000000000046': FType_Word6,
    # EXCEL
    '00020820-0000-0000-C000-000000000046': FType_Excel97,
    '00020810-0000-0000-C000-000000000046': FType_Excel5,
    # POWERPOINT
    '64818D10-4F9B-11CF-86EA-00AA00B929E8': FType_Powerpoint97,
    # MSI
    '000C1084-0000-0000-C000-000000000046': FType_MSI,
}

openxml_ftypes = {
    # mapping from content-type of main part to FType classes:
    # WORD
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml': FType_Word2007,
    'application/vnd.ms-word.document.macroEnabled.main+xml': FType_Word2007_Macro,
    'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml': FType_Word2007_Template,
    'application/vnd.ms-word.template.macroEnabledTemplate.main+xml': FType_Word2007_Template_Macro,
    # EXCEL
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml': FType_Excel2007_XLSX,
    'application/vnd.ms-excel.sheet.macroEnabled.main+xml': FType_Excel2007_XLSM,
    'application/vnd.ms-excel.sheet.binary.macroEnabled.main': FType_Excel2007_XLSB,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml': FType_Excel2007_Template,
    'application/vnd.ms-excel.template.macroEnabled.main+xml': FType_Excel2007_Template_Macro,
    'application/vnd.ms-excel.addin.macroEnabled.main+xml': FType_Excel2007_Addin_Macro,
    # POWERPOINT
    'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml': FType_Powerpoint2007_Presentation, #PPTX
    'application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml': FType_Powerpoint2007_Slideshow, #PPSX
    'application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml': FType_Powerpoint2007_Macro, #PPTM
    'application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml': FType_Powerpoint2007_Slideshow_Macro, #PPSM

    # TODO: add missing PowerPoint formats:
    # PPAM – PowerPoint Add-in Open Office XML File Format. Mime type is application/vnd.ms-powerpoint.addin.macroEnabled.12.
    # POTX – PowerPoint Template Open Office XML File Format. Mime type is application/vnd.openxmlformats-officedocument.presentationml.template.
    # POTM – PowerPoint Macro-Enabled Template Open Office XML File Format. Mime type is application/vnd.ms-powerpoint.template.macroEnabled.12.

    # XPS
    'application/vnd.ms-package.xps-fixeddocumentsequence+xml': FType_XPS,
    #TODO: Add MSIX
}


class FType_EXE_PE (FType_Base):
    filetype = FTYPE.EXE_PE
    container = CONTAINER.BINARY
    application = APP.WINDOWS
    name = "Windows PE Executable or DLL"
    longname = "Windows Portable Executable or DLL (EXE,DLL)"
    extensions = ('exe', 'dll', 'sys', 'scr')  # TODO: add more from https://en.wikipedia.org/wiki/Portable_Executable
    content_types = ('application/vnd.microsoft.portable-executable',)
    PUID = 'fmt/899'

    @classmethod
    def recognize(cls, ftg):
        return True if ftg.data.startswith(b'MZ') else False
        # TODO: make this more accurate by checking the PE header, e.g. using pefile or directly

class FileTypeGuesser(object):
    """
    A class to guess the type of a file, focused on MS Office, RTF and ZIP.
    """

    def __init__(self, filepath=None, data=None):
        self.filepath = filepath
        self.data = data
        self.container = None
        self.application = None
        self.filetype = None
        self.ftype = FType_Unknown  # FType class
        self.data_bytesio = None
        # For OLE:
        self.olefile = None
        self.root_clsid = None
        self.root_clsid_name = None
        # For ZIP:
        self.zipfile = None
        # For OpenXML:
        self.root_rels = None
        # self.main_part_xmltag = None
        self.main_part_content_type = None
        # For XML:
        self.root_xmltag = None
        self.xmlroot = None

        if filepath is None and data is None:
            raise ValueError('FileTypeGuesser requires either a file path or file data, or both')
        if data is None:
            with open(filepath, 'rb') as f:
                self.data = f.read()
        self.data_bytesio = io.BytesIO(self.data)

        # Identify the main container type:
        for ftype in (FType_RTF, FType_Generic_OLE, FType_Generic_Zip, FType_OneNote, FType_PNG):
            if ftype.recognize(self):
                self.ftype = ftype
                break
        self.container = self.ftype.container
        self.filetype = self.ftype.filetype
        self.application = self.ftype.application

        # OLE file types:
        if self.container == CONTAINER.OLE:
            # for ftype in (FType_Word97, FType_Word6, FType_Excel97, FType_Excel5):
            #     if ftype.recognize(self):
            #         self.ftype = ftype
            #         break
            ft = clsid_ftypes.get(self.root_clsid, None)
            if ft is not None:
                self.ftype = ft

        # OpenXML file types:
        if self.container == CONTAINER.ZIP:
            if FType_Generic_OpenXML.recognize(self):
                self.ftype = FType_Generic_OpenXML
                ft = openxml_ftypes.get(self.main_part_content_type, None)
                if ft is not None:
                    self.ftype = ft

        # TODO: use a mapping from magic to file types
        if self.container == CONTAINER.UNKNOWN:
            if FType_EXE_PE.recognize(self):
                self.ftype = FType_EXE_PE

        self.container = self.ftype.container
        self.filetype = self.ftype.filetype
        self.application = self.ftype.application

    def __str__(self):
        """Give a short string representation of this object."""
        return '[FileTypeGuesser for {0}: {1} from {2} in {3}]'.format(
            "data" if self.filepath is None
            else os.path.basename(self.filepath),
            self.filetype, self.application, self.container)

    def close(self):
        """
        This method must be called at the end of processing
        """
        # TODO: only close self.olefile if it was opened by ftguess
        if self.zipfile is not None:
            self.zipfile.close()

    def is_ole(self):
        """
        Shortcut to check if the container is OLE
        :return: bool
        """
        return issubclass(self.ftype, FType_Generic_OLE) or self.container == CONTAINER.OLE

    def is_openxml(self):
        """
        Shortcut to check if the container is OpenXML
        :return: bool
        """
        return issubclass(self.ftype, FType_Generic_OpenXML) or self.container == CONTAINER.OpenXML

    def is_word(self):
        """
        Shortcut to check if a file is an Excel workbook, template or add-in
        :return: bool
        """
        return issubclass(self.ftype, FType_Word)

    def is_excel(self):
        """
        Shortcut to check if a file is an Excel workbook, template or add-in
        :return: bool
        """
        return issubclass(self.ftype, FType_Excel)

    def is_powerpoint(self):
        """
        Shortcut to check if a file is Powerpoint file of any kind
        :return: bool
        """
        return issubclass(self.ftype, FType_Powerpoint)


# === FUNCTIONS ==============================================================

def ftype_guess(filepath=None, data=None):
    return FileTypeGuesser(filepath, data)

def process_file(container, filename, data):
    print('File       : %s' % filename)
    ftg = ftype_guess(filepath=filename, data=data)
    print('File Type  : %s' % ftg.ftype.name)
    print('Description: %s' % ftg.ftype.longname)
    print('Application: %s' % ftg.ftype.application)
    print('Container  : %s' % ftg.container)
    if ftg.root_clsid is not None:
        print('Root CLSID : %s - %s' % (ftg.root_clsid, ftg.root_clsid_name))
    print('Content-type(s) : %s' % ','.join(ftg.ftype.content_types))
    print('PUID       : %s' % ftg.ftype.PUID)
    print()


#=== MAIN =================================================================

def main():
    # print banner with version
    python_version = '%d.%d.%d' % sys.version_info[0:3]
    print ('ftguess %s on Python %s - http://decalage.info/python/oletools' %
           (__version__, python_version))
    print ('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print ('Please report any issue at https://github.com/decalage2/oletools/issues')
    print ('')

    DEFAULT_LOG_LEVEL = "warning" # Default log level
    LOG_LEVELS = {
        'debug':    logging.DEBUG,
        'info':     logging.INFO,
        'warning':  logging.WARNING,
        'error':    logging.ERROR,
        'critical': logging.CRITICAL
        }

    usage = 'usage: %prog [options] <filename> [filename2 ...]'
    parser = optparse.OptionParser(usage=usage)
    # parser.add_option('-c', '--csv', dest='csv',
    #     help='export results to a CSV file')
    parser.add_option("-r", action="store_true", dest="recursive",
        help='find files recursively in subdirectories.')
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
        help='if the file is a zip archive, open first file from it, using the provided password')
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
    enable_logging()

    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
        zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        process_file(container, filename, data)


if __name__ == '__main__':
    main()
