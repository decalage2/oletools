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

# oleid is copyright (c) 2012-2021, Philippe Lagadec (http://www.decalage.info)
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

__version__ = '0.56.2'


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

import argparse, sys, re, zlib, struct, os
from os.path import dirname, abspath

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

from oletools.thirdparty.prettytable import prettytable
from oletools import crypto



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
                 description=None):
        self.id = _id
        self.value = value
        self.type = _type
        self.name = name
        if name == None:
            self.name = _id
        self.description = description


class OleID(object):
    """
    Summary of information about an OLE file

    Call :py:meth:`OleID.check` to gather all info on a given file or run one
    of the `check_` functions to just get a specific piece of info.
    """

    def __init__(self, input_file):
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
        if isinstance(input_file, olefile.OleFileIO):
            self.ole = input_file
            self.filename = None
        else:
            self.filename = input_file
            self.ole = None
        self.indicators = []
        self.suminfo_data = None

    def check(self):
        """
        Open file and run all checks on it.

        :returns: list of all :py:class:`Indicator`s created
        """
        # check if it is actually an OLE file:
        oleformat = Indicator('ole_format', True, name='OLE format')
        self.indicators.append(oleformat)
        if self.ole:
            oleformat.value = True
        elif not olefile.isOleFile(self.filename):
            oleformat.value = False
            return self.indicators
        else:
            # parse file:
            self.ole = olefile.OleFileIO(self.filename)
        # checks:
        self.check_properties()
        self.check_encrypted()
        self.check_word()
        self.check_excel()
        self.check_powerpoint()
        self.check_visio()
        self.check_object_pool()
        self.check_flash()
        self.ole.close()
        return self.indicators

    def check_properties(self):
        """
        Read summary information required for other check_* functions

        :returns: 2 :py:class:`Indicator`s (for presence of summary info and
                    application name) or None if file was not opened
        """
        suminfo = Indicator('has_suminfo', False,
                            name='Has SummaryInformation stream')
        self.indicators.append(suminfo)
        appname = Indicator('appname', 'unknown', _type=str,
                            name='Application name')
        self.indicators.append(appname)
        if not self.ole:
            return None, None
        self.suminfo_data = {}
        # check stream SummaryInformation (not present e.g. in encrypted ppt)
        if self.ole.exists("\x05SummaryInformation"):
            suminfo.value = True
            self.suminfo_data = self.ole.getproperties("\x05SummaryInformation")
            # check application name:
            appname.value = self.suminfo_data.get(0x12, 'unknown')
        return suminfo, appname

    def get_indicator(self, indicator_id):
        """Helper function: returns an indicator if present (or None)"""
        result = [indicator for indicator in self.indicators
                  if indicator.id == indicator_id]
        if result:
            return result[0]
        else:
            return None

    def check_encrypted(self):
        """
        Check whether this file is encrypted.

        Might call check_properties.

        :returns: :py:class:`Indicator` for encryption or None if file was not
                  opened
        """
        # we keep the pointer to the indicator, can be modified by other checks:
        encrypted = Indicator('encrypted', False, name='Encrypted')
        self.indicators.append(encrypted)
        if not self.ole:
            return None
        encrypted.value = crypto.is_encrypted(self.ole)
        return encrypted

    def check_word(self):
        """
        Check whether this file is a word document

        If this finds evidence of encryption, will correct/add encryption
        indicator.

        :returns: 2 :py:class:`Indicator`s (for word and vba_macro) or None if
                  file was not opened
        """
        word = Indicator(
            'word', False, name='Word Document',
            description='Contains a WordDocument stream, very likely to be a '
                        'Microsoft Word Document.')
        self.indicators.append(word)
        macros = Indicator('vba_macros', False, name='VBA Macros')
        self.indicators.append(macros)
        if not self.ole:
            return None, None
        if self.ole.exists('WordDocument'):
            word.value = True

            # check for VBA macros:
            if self.ole.exists('Macros'):
                macros.value = True
        return word, macros

    def check_excel(self):
        """
        Check whether this file is an excel workbook.

        If this finds macros, will add/correct macro indicator.

        see also: :py:func:`xls_parser.is_xls`

        :returns: :py:class:`Indicator` for excel or (None, None) if file was
                  not opened
        """
        excel = Indicator(
            'excel', False, name='Excel Workbook',
            description='Contains a Workbook or Book stream, very likely to be '
                        'a Microsoft Excel Workbook.')
        self.indicators.append(excel)
        if not self.ole:
            return None
        #self.macros = Indicator('vba_macros', False, name='VBA Macros')
        #self.indicators.append(self.macros)
        if self.ole.exists('Workbook') or self.ole.exists('Book'):
            excel.value = True
            # check for VBA macros:
            if self.ole.exists('_VBA_PROJECT_CUR'):
                # correct macro indicator if present or add one
                macro_ind = self.get_indicator('vba_macros')
                if macro_ind:
                    macro_ind.value = True
                else:
                    macros = Indicator('vba_macros', True, name='VBA Macros')
                    self.indicators.append(macros)
        return excel

    def check_powerpoint(self):
        """
        Check whether this file is a powerpoint presentation

        see also: :py:func:`ppt_record_parser.is_ppt`

        :returns: :py:class:`Indicator` for whether this is a powerpoint
                  presentation or not or None if file was not opened
        """
        ppt = Indicator(
            'ppt', False, name='PowerPoint Presentation',
            description='Contains a PowerPoint Document stream, very likely to '
                        'be a Microsoft PowerPoint Presentation.')
        self.indicators.append(ppt)
        if not self.ole:
            return None
        if self.ole.exists('PowerPoint Document'):
            ppt.value = True
        return ppt

    def check_visio(self):
        """Check whether this file is a visio drawing"""
        visio = Indicator(
            'visio', False, name='Visio Drawing',
            description='Contains a VisioDocument stream, very likely to be a '
                        'Microsoft Visio Drawing.')
        self.indicators.append(visio)
        if not self.ole:
            return None
        if self.ole.exists('VisioDocument'):
            visio.value = True
        return visio

    def check_object_pool(self):
        """
        Check whether this file contains an ObjectPool stream.

        Such a stream would be a strong indicator for embedded objects or files.

        :returns: :py:class:`Indicator` for ObjectPool stream or None if file
                  was not opened
        """
        objpool = Indicator(
            'ObjectPool', False, name='ObjectPool',
            description='Contains an ObjectPool stream, very likely to contain '
                        'embedded OLE objects or files.')
        self.indicators.append(objpool)
        if not self.ole:
            return None
        if self.ole.exists('ObjectPool'):
            objpool.value = True
        return objpool

    def check_flash(self):
        """
        Check whether this file contains flash objects

        :returns: :py:class:`Indicator` for count of flash objects or None if
                  file was not opened
        """
        flash = Indicator(
            'flash', 0, _type=int, name='Flash objects',
            description='Number of embedded Flash objects (SWF files) detected '
                        'in OLE streams. Not 100% accurate, there may be false '
                        'positives.')
        self.indicators.append(flash)
        if not self.ole:
            return None
        for stream in self.ole.listdir():
            data = self.ole.openstream(stream).read()
            found = detect_flash(data)
            # just add to the count of Flash objects:
            flash.value += len(found)
            #print stream, found
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

    for filename in args.input:
        print('Filename:', filename)
        oleid = OleID(filename)
        indicators = oleid.check()

        #TODO: add description
        #TODO: highlight suspicious indicators
        table = prettytable.PrettyTable(['Indicator', 'Value'])
        table.align = 'l'
        table.max_width = 39
        table.border = False

        for indicator in indicators:
            #print '%s: %s' % (indicator.name, indicator.value)
            table.add_row((indicator.name, indicator.value))

        print(table)
        print('')

if __name__ == '__main__':
    main()
