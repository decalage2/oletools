#!/usr/bin/env python
"""
oleid.py

oleid is a script to analyze OLE files such as MS Office documents (e.g. Word,
Excel), to detect specific characteristics that could potentially indicate that
the file is suspicious or malicious, in terms of security (e.g. malware).
For example it can detect VBA macros, embedded Flash objects, fragmentation.
The results is displayed as ascii table (but could be returned or printed in
other formats like CSV, XML or JSON in future).

oleid project website: http://www.decalage.info/python/oleid

oleid is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE =================================================================

# oleid is copyright (c) 2012-2022, Philippe Lagadec (http://www.decalage.info)
# All rights reserved.
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  * Redistributions of source code must retain the above copyright notice, this
#    list of conditions and the following disclaimer.
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

# To improve Python 2+3 compatibility:
from __future__ import print_function
#from __future__ import absolute_import

#------------------------------------------------------------------------------
# CHANGELOG:
# 2012-10-29 v0.01 PL: - first version
# 2014-11-29 v0.02 PL: - use olefile instead of OleFileIO_PL
#                      - improved usage display with -h
# 2014-11-30 v0.03 PL: - improved output with prettytable
# 2016-10-25 v0.50 PL: - fixed print and bytes strings for Python 3
# 2016-12-12 v0.51 PL: - fixed relative imports for Python 3 (issue #115)
# 2017-04-26       PL: - fixed absolute imports (issue #141)
# 2017-09-01       SA: - detect OpenXML encryption
# 2018-09-11 v0.54 PL: - olefile is now a dependency
# 2018-10-19       CH: - accept olefile as well as filename, return Indicators,
#                        improve encryption detection for ppt
# 2021-05-07 v0.56.2 MN: - fixed bug in check_excel (issue #584, PR #585)

__version__ = '0.60.1'


#------------------------------------------------------------------------------
# TODO:
# + extract relevant metadata: codepage, author, application, timestamps, etc
# - detect RTF and OpenXML
# - fragmentation
# - OLE package
# - entropy
# - detect PE header?
# - detect NOPs?
# - list type of each object in object pool?
# - criticality for each indicator?: info, low, medium, high
# - support wildcards with glob?
# - verbose option
# - csv, xml output


#=== IMPORTS =================================================================

import argparse, sys, re, zlib, struct, os, io

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

from oletools.thirdparty.tablestream import tablestream
from oletools import crypto, ftguess, olevba, mraptor, oleobj, ooxml
from oletools.common.log_helper import log_helper
from oletools.common.codepages import get_codepage_name

# === LOGGING =================================================================

log = log_helper.get_or_create_silent_logger('oleid')

# === CONSTANTS ===============================================================

class RISK(object):
    """
    Constants for risk levels
    """
    HIGH = 'HIGH'
    MEDIUM = 'Medium'
    LOW = 'low'
    NONE = 'none'
    INFO = 'info'
    UNKNOWN = 'Unknown'
    ERROR = 'Error'  # if a check triggered an unexpected error

risk_color = {
    RISK.HIGH: 'red',
    RISK.MEDIUM: 'yellow',
    RISK.LOW: 'white',
    RISK.NONE: 'green',
    RISK.INFO: 'cyan',
    RISK.UNKNOWN: None,
    RISK.ERROR: None
}

#=== FUNCTIONS ===============================================================

def detect_flash(data):
    """
    Detect Flash objects (SWF files) within a binary string of data
    return a list of (start_index, length, compressed) tuples, or [] if nothing
    found.

    Code inspired from xxxswf.py by Alexander Hanel (but significantly reworked)
    http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html
    """
    #TODO: report
    found = []
    for match in re.finditer(b'CWS|FWS', data):
        start = match.start()
        if start+8 > len(data):
            # header size larger than remaining data, this is not a SWF
            continue
        #TODO: one struct.unpack should be simpler
        # Read Header
        header = data[start:start+3]
        # Read Version
        ver = struct.unpack('<b', data[start+3:start+4])[0]
        # Error check for version above 20
        #TODO: is this accurate? (check SWF specifications)
        if ver > 20:
            continue
        # Read SWF Size
        size = struct.unpack('<i', data[start+4:start+8])[0]
        if start+size > len(data) or size < 1024:
            # declared size larger than remaining data, this is not a SWF
            # or declared size too small for a usual SWF
            continue
        # Read SWF into buffer. If compressed read uncompressed size.
        swf = data[start:start+size]
        compressed = False
        if b'CWS' in header:
            compressed = True
            # compressed SWF: data after header (8 bytes) until the end is
            # compressed with zlib. Attempt to decompress it to check if it is
            # valid
            compressed_data = swf[8:]
            try:
                zlib.decompress(compressed_data)
            except Exception:
                continue
        # else we don't check anything at this stage, we only assume it is a
        # valid SWF. So there might be false positives for uncompressed SWF.
        found.append((start, size, compressed))
        #print 'Found SWF start=%x, length=%d' % (start, size)
    return found


#=== CLASSES =================================================================

class Indicator(object):
    """
    Piece of information of an :py:class:`OleID` object.

    Contains an ID, value, type, name and description. No other functionality.
    """

    def __init__(self, _id, value=None, _type=bool, name=None,
                 description=None, risk=RISK.UNKNOWN, hide_if_false=True):
        self.id = _id
        self.value = value
        self.type = _type
        self.name = name
        if name == None:
            self.name = _id
        self.description = description
        self.risk = risk
        self.hide_if_false = hide_if_false


class OleID(object):
    """
    Summary of information about an OLE file (and a few other MS Office formats)

    Call :py:meth:`OleID.check` to gather all info on a given file or run one
    of the `check_` functions to just get a specific piece of info.
    """

    def __init__(self, filename=None, data=None):
        """
        Create an OleID object

        This does not run any checks yet nor open the file.

        Can either give just a filename (as str), so OleID will check whether
        that is a valid OLE file and create a :py:class:`olefile.OleFileIO`
        object for it. Or you can give an already opened
        :py:class:`olefile.OleFileIO` as argument to avoid re-opening (e.g. if
        called from other oletools).

        If filename is given, only :py:meth:`OleID.check` opens the file. Other
        functions will return None
        """
        if filename is None and data is None:
            raise ValueError('OleID requires either a file path or file data, or both')
        self.file_on_disk = False  # True = file on disk / False = file in memory
        if data is None:
            self.file_on_disk = True  # useful for some check that don't work in memory
            with open(filename, 'rb') as f:
                self.data = f.read()
        else:
            self.data = data
        self.data_bytesio = io.BytesIO(self.data)
        if isinstance(filename, olefile.OleFileIO):
            self.ole = filename
            self.filename = None
        else:
            self.filename = filename
            self.ole = None
        self.indicators = []
        self.suminfo_data = None

    def get_indicator(self, indicator_id):
        """Helper function: returns an indicator if present (or None)"""
        result = [indicator for indicator in self.indicators
                  if indicator.id == indicator_id]
        if result:
            return result[0]
        else:
            return None

    def check(self):
        """
        Open file and run all checks on it.

        :returns: list of all :py:class:`Indicator`s created
        """
        self.ftg = ftguess.FileTypeGuesser(filepath=self.filename, data=self.data)
        ftype = self.ftg.ftype
        # if it's an unrecognized OLE file, display the root CLSID in description:
        if self.ftg.filetype == ftguess.FTYPE.GENERIC_OLE:
            description = 'Unrecognized OLE file. Root CLSID: {} - {}'.format(
                self.ftg.root_clsid, self.ftg.root_clsid_name)
        else:
            description = ''
        ft = Indicator('ftype', value=ftype.longname, _type=str, name='File format', risk=RISK.INFO,
                       description=description)
        self.indicators.append(ft)
        ct = Indicator('container', value=ftype.container, _type=str, name='Container format', risk=RISK.INFO,
                       description='Container type')
        self.indicators.append(ct)

        # check if it is actually an OLE file:
        if self.ftg.container == ftguess.CONTAINER.OLE:
            # reuse olefile already opened by ftguess
            self.ole = self.ftg.olefile
        # oleformat = Indicator('ole_format', True, name='OLE format')
        # self.indicators.append(oleformat)
        # if self.ole:
        #     oleformat.value = True
        # elif not olefile.isOleFile(self.filename):
        #     oleformat.value = False
        #     return self.indicators
        # else:
        #     # parse file:
        #     self.ole = olefile.OleFileIO(self.filename)

        # checks:
        # TODO: add try/except around each check
        self.check_properties()
        self.check_encrypted()
        self.check_macros()
        self.check_external_relationships()
        self.check_object_pool()
        self.check_flash()
        if self.ole is not None:
            self.ole.close()
        return self.indicators

    def check_properties(self):
        """
        Read summary information required for other check_* functions

        :returns: 2 :py:class:`Indicator`s (for presence of summary info and
                    application name) or None if file was not opened
        """
        if not self.ole:
            return None
        meta = self.ole.get_metadata()
        appname = Indicator('appname', meta.creating_application, _type=str,
                            name='Application name', description='Application name declared in properties',
                            risk=RISK.INFO)
        self.indicators.append(appname)
        codepage_name = None
        if meta.codepage is not None:
            codepage_name = '{}: {}'.format(meta.codepage, get_codepage_name(meta.codepage))
        codepage = Indicator('codepage', codepage_name, _type=str,
                      name='Properties code page', description='Code page used for properties',
                      risk=RISK.INFO)
        self.indicators.append(codepage)
        author = Indicator('author', meta.author, _type=str,
                      name='Author', description='Author declared in properties',
                      risk=RISK.INFO)
        self.indicators.append(author)
        return appname, codepage, author

    def check_encrypted(self):
        """
        Check whether this file is encrypted.

        :returns: :py:class:`Indicator` for encryption or None if file was not
                  opened
        """
        # we keep the pointer to the indicator, can be modified by other checks:
        encrypted = Indicator('encrypted', False, name='Encrypted',
                              risk=RISK.NONE,
                              description='The file is not encrypted',
                              hide_if_false=False)
        self.indicators.append(encrypted)
        # Only OLE files can be encrypted (OpenXML files are encrypted in an OLE container):
        if not self.ole:
            return None
        try:
            if crypto.is_encrypted(self.ole):
                encrypted.value = True
                encrypted.risk = RISK.LOW
                encrypted.description = 'The file is encrypted. It may be decrypted with msoffcrypto-tool'
        except Exception as exception:
            # msoffcrypto-tool can trigger exceptions, such as "Unknown file format" for Excel 5.0/95
            encrypted.value = 'Error'
            encrypted.risk = RISK.ERROR
            encrypted.description = 'msoffcrypto-tool raised an error when checking if the file is encrypted: {}'.format(exception)
        return encrypted

    def check_external_relationships(self):
        """
        Check whether this file has external relationships (remote template, OLE object, etc).

        :returns: :py:class:`Indicator`
        """
        ext_rels = Indicator('ext_rels', 0, name='External Relationships', _type=int,
                              risk=RISK.NONE,
                              description='External relationships such as remote templates, remote OLE objects, etc',
                              hide_if_false=False)
        self.indicators.append(ext_rels)
        # this check only works for OpenXML files
        if not self.ftg.is_openxml():
            return ext_rels
        # to collect relationship types:
        rel_types = set()
        # open an XmlParser, using a BytesIO instead of filename (to work in memory)
        xmlparser = ooxml.XmlParser(self.data_bytesio)
        for rel_type, target in oleobj.find_external_relationships(xmlparser):
            log.debug('External relationship: type={} target={}'.format(rel_type, target))
            rel_types.add(rel_type)
            ext_rels.value += 1
        if ext_rels.value > 0:
            ext_rels.description = 'External relationships found: {} - use oleobj for details'.format(
                ', '.join(rel_types))
            ext_rels.risk = RISK.HIGH
        return ext_rels

    def check_object_pool(self):
        """
        Check whether this file contains an ObjectPool stream.

        Such a stream would be a strong indicator for embedded objects or files.

        :returns: :py:class:`Indicator` for ObjectPool stream or None if file
                  was not opened
        """
        # TODO: replace this by a call to oleobj + add support for OpenXML
        objpool = Indicator(
            'ObjectPool', False, name='ObjectPool',
            description='Contains an ObjectPool stream, very likely to contain '
                        'embedded OLE objects or files. Use oleobj to check it.',
            risk=RISK.NONE)
        self.indicators.append(objpool)
        if not self.ole:
            return None
        if self.ole.exists('ObjectPool'):
            objpool.value = True
            objpool.risk = RISK.LOW
            # TODO: set risk to medium for OLE package if not executable
            # TODO: set risk to high for Package executable or object with CVE in CLSID
        return objpool

    def check_macros(self):
        """
        Check whether this file contains macros (VBA and XLM/Excel 4).

        :returns: :py:class:`Indicator`
        """
        vba_indicator = Indicator(_id='vba', value='No', _type=str, name='VBA Macros',
                                  description='This file does not contain VBA macros.',
                                  risk=RISK.NONE, hide_if_false=False)
        self.indicators.append(vba_indicator)
        xlm_indicator = Indicator(_id='xlm', value='No', _type=str, name='XLM Macros',
                                  description='This file does not contain Excel 4/XLM macros.',
                                  risk=RISK.NONE, hide_if_false=False)
        self.indicators.append(xlm_indicator)
        if self.ftg.filetype == ftguess.FTYPE.RTF:
            # For RTF we don't call olevba otherwise it triggers an error
            vba_indicator.description = 'RTF files cannot contain VBA macros'
            xlm_indicator.description = 'RTF files cannot contain XLM macros'
            return vba_indicator, xlm_indicator
        vba_parser = None  # flag in case olevba fails
        try:
            vba_parser = olevba.VBA_Parser(filename=self.filename, data=self.data)
            if vba_parser.detect_vba_macros():
                vba_indicator.value = 'Yes'
                vba_indicator.risk = RISK.MEDIUM
                vba_indicator.description = 'This file contains VBA macros. No suspicious keyword was found. Use olevba and mraptor for more info.'
                # check code with mraptor
                vba_code = vba_parser.get_vba_code_all_modules()
                m = mraptor.MacroRaptor(vba_code)
                m.scan()
                if m.suspicious:
                    vba_indicator.value = 'Yes, suspicious'
                    vba_indicator.risk = RISK.HIGH
                    vba_indicator.description = 'This file contains VBA macros. Suspicious keywords were found. Use olevba and mraptor for more info.'
        except Exception as e:
            vba_indicator.risk = RISK.ERROR
            vba_indicator.value = 'Error'
            vba_indicator.description = 'Error while checking VBA macros: %s' % str(e)
        finally:
            if vba_parser is not None:
                vba_parser.close()
            vba_parser = None
        # Check XLM macros only for Excel file types:
        if self.ftg.is_excel():
            # TODO: for now XLM detection only works for files on disk... So we need to reload VBA_Parser from the filename
            #       To be improved once XLMMacroDeobfuscator can work on files in memory
            if self.file_on_disk:
                try:
                    vba_parser = olevba.VBA_Parser(filename=self.filename)
                    if vba_parser.detect_xlm_macros():
                        xlm_indicator.value = 'Yes'
                        xlm_indicator.risk = RISK.MEDIUM
                        xlm_indicator.description = 'This file contains XLM macros. Use olevba to analyse them.'
                except Exception as e:
                    xlm_indicator.risk = RISK.ERROR
                    xlm_indicator.value = 'Error'
                    xlm_indicator.description = 'Error while checking XLM macros: %s' % str(e)
                finally:
                    if vba_parser is not None:
                        vba_parser.close()
            else:
                xlm_indicator.risk = RISK.UNKNOWN
                xlm_indicator.value = 'Unknown'
                xlm_indicator.description = 'For now, XLM macros can only be detected for files on disk, not in memory'

        return vba_indicator, xlm_indicator

    def check_flash(self):
        """
        Check whether this file contains flash objects

        :returns: :py:class:`Indicator` for count of flash objects or None if
                  file was not opened
        """
        # TODO: add support for RTF and OpenXML formats
        flash = Indicator(
            'flash', 0, _type=int, name='Flash objects',
            description='Number of embedded Flash objects (SWF files) detected '
                        'in OLE streams. Not 100% accurate, there may be false '
                        'positives.',
            risk=RISK.NONE)
        self.indicators.append(flash)
        if not self.ole:
            return None
        for stream in self.ole.listdir():
            data = self.ole.openstream(stream).read()
            found = detect_flash(data)
            # just add to the count of Flash objects:
            flash.value += len(found)
            #print stream, found
        if flash.value > 0:
            flash.risk = RISK.MEDIUM
        return flash


#=== MAIN =================================================================

def main():
    """Called when running this file as script. Shows all info on input file."""
    # print banner with version
    print('oleid %s - http://decalage.info/oletools' % __version__)
    print('THIS IS WORK IN PROGRESS - Check updates regularly!')
    print('Please report any issue at '
          'https://github.com/decalage2/oletools/issues')
    print('')

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input', type=str, nargs='*', metavar='FILE',
                        help='Name of files to process')
    # parser.add_argument('-o', '--ole', action='store_true', dest='ole',
    #                   help='Parse an OLE file (e.g. Word, Excel) to look for '
    #                        'SWF in each stream')

    args = parser.parse_args()

    # Print help if no argurments are passed
    if len(args.input) == 0:
        parser.print_help()
        return

    log_helper.enable_logging()

    for filename in args.input:
        print('Filename:', filename)
        oleid = OleID(filename)
        indicators = oleid.check()

        table = tablestream.TableStream([20, 20, 10, 26],
                                        header_row=['Indicator', 'Value', 'Risk', 'Description'],
                                        style=tablestream.TableStyleSlimSep)
        for indicator in indicators:
            if not (indicator.hide_if_false and not indicator.value):
                #print '%s: %s' % (indicator.name, indicator.value)
                color = risk_color.get(indicator.risk, None)
                table.write_row((indicator.name, indicator.value, indicator.risk, indicator.description),
                                colors=(color, color, color, None))
        table.close()

if __name__ == '__main__':
    main()
