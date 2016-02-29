#!/usr/bin/env python
"""
mraptor.py - MacroRaptor

MacroRaptor is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to detect malicious macros.

Supported formats:
- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 2007+ (.pptm, .ppsm)
- Word 2003 XML (.xml)
- Word/Excel Single File Web Page / MHTML (.mht)

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

MacroRaptor is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

# === LICENSE ==================================================================

# MacroRaptor is copyright (c) 2016 Philippe Lagadec (http://www.decalage.info)
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
# 2016-02-23 v0.01 PL: - first version
# 2016-02-29 v0.02 PL: - added Workbook_Activate, FileSaveAs

__version__ = '0.01'

#------------------------------------------------------------------------------
# TODO:


#--- IMPORTS ------------------------------------------------------------------

import sys, logging, optparse, re

from thirdparty.xglob import xglob
from thirdparty.tablestream import tablestream

import olevba

# === LOGGING =================================================================

# a global logger object used for debugging:
log = olevba.get_logger('mraptor')


#--- CONSTANTS ----------------------------------------------------------------

# URL and message to report issues:
URL_ISSUES = 'https://bitbucket.org/decalage/oletools/issues'
MSG_ISSUES = 'Please report this issue on %s' % URL_ISSUES

# 'AutoExec', 'AutoOpen', 'Auto_Open', 'AutoClose', 'Auto_Close', 'AutoNew', 'AutoExit',
# 'Document_Open', 'DocumentOpen',
# 'Document_Close', 'DocumentBeforeClose',
# 'DocumentChange','Document_New',
# 'NewDocument'
# 'Workbook_Open', 'Workbook_Close',

# TODO: check if line also contains Sub or Function
re_autoexec = re.compile(r'(?i)\b(?:Auto(?:Exec|_?Open|_?Close|Exit|New)' +
                         r'|Document(?:_?Open|_Close|BeforeClose|Change|_New)' +
                         r'|NewDocument|Workbook(?:_Open|_Activate|_Close))\b')

# MS-VBAL 5.4.5.1 Open Statement:
RE_OPEN_WRITE = r'(?:\bOpen\b[^\n]+\b(?:Write|Append|Binary|Output|Random)\b)'

re_write = re.compile(r'(?i)\b(?:FileCopy|CopyFile|Kill|CreateTextFile|'
    + r'VirtualAlloc|RtlMoveMemory|URLDownloadToFileA?|AltStartupPath|'
    + r'ADODB\.Stream|WriteText|SaveToFile|SaveAs|SaveAsRTF|FileSaveAs|MkDir|RmDir|SaveSetting|SetAttr)\b|' + RE_OPEN_WRITE)

# MS-VBAL 5.2.3.5 External Procedure Declaration
RE_DECLARE_LIB = r'(?:\bDeclare\b[^\n]+\bLib\b)'

re_execute = re.compile(r'(?i)\b(?:Shell|CreateObject|GetObject|SendKeys|'
    + r'MacScript|FollowHyperlink|CreateThread|ShellExecute)\b|' + RE_DECLARE_LIB)

# short tag to display file types in triage mode:
TYPE2TAG = {
    olevba.TYPE_OLE: 'OLE',
    olevba.TYPE_OpenXML: 'OpX',
    olevba.TYPE_Word2003_XML: 'XML',
    olevba.TYPE_MHTML: 'MHT',
    olevba.TYPE_TEXT: 'TXT',
}


# === CLASSES =================================================================

class MacroRaptor(object):
    """
    class to scan VBA macro code to detect if it is malicious
    """
    def __init__(self, vba_code):
        """
        MacroRaptor constructor
        :param vba_code: string containing the VBA macro code
        """
        # TODO: collapse long lines first
        self.vba_code = vba_code
        self.autoexec = False
        self.write = False
        self.execute = False
        self.flags = ''
        self.suspicious = False
        self.autoexec_match = None
        self.write_match = None
        self.execute_match = None
        self.matches = []

    def scan(self):
        """
        Scan the VBA macro code to detect if it is malicious
        :return:
        """
        m = re_autoexec.search(self.vba_code)
        if m is not None:
            self.autoexec = True
            self.autoexec_match = m.group()
            self.matches.append(m.group())
        m = re_write.search(self.vba_code)
        if m is not None:
            self.write = True
            self.write_match = m.group()
            self.matches.append(m.group())
        m = re_execute.search(self.vba_code)
        if m is not None:
            self.execute = True
            self.execute_match = m.group()
            self.matches.append(m.group())
        if self.autoexec and (self.execute or self.write):
            self.suspicious = True

    def get_flags(self):
        flags = ''
        flags += 'A' if self.autoexec else '-'
        flags += 'W' if self.write else '-'
        flags += 'X' if self.execute else '-'
        return flags


# === MAIN ====================================================================

def main():
    """
    Main function, called when olevba is run from the command line
    """
    global log
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
    parser.add_option("-r", action="store_true", dest="recursive",
                      help='find files recursively in subdirectories.')
    parser.add_option("-z", "--zip", dest='zip_password', type='str', default=None,
                      help='if the file is a zip archive, open all files from it, using the provided password (requires Python 2.6+)')
    parser.add_option("-f", "--zipfname", dest='zip_fname', type='str', default='*',
                      help='if the file is a zip archive, file(s) to be opened within the zip. Wildcards * and ? are supported. (default:*)')
    parser.add_option('-l', '--loglevel', dest="loglevel", action="store", default=DEFAULT_LOG_LEVEL,
                            help="logging level debug/info/warning/error/critical (default=%default)")
    parser.add_option("-m", '--matches', action="store_true", dest="show_matches",
                      help='Show matched strings.')

    # TODO: add logfile option

    (options, args) = parser.parse_args()

    # Print help if no arguments are passed
    if len(args) == 0:
        print __doc__
        parser.print_help()
        sys.exit()

    # print banner with version
    print 'MacroRaptor %s - http://decalage.info/python/oletools' % __version__
    print 'This is work in progress, please report issues at %s' % URL_ISSUES

    logging.basicConfig(level=LOG_LEVELS[options.loglevel], format='%(levelname)-8s %(message)s')
    # enable logging in the modules:
    log.setLevel(logging.NOTSET)

    t = tablestream.TableStream(style=tablestream.TableStyleSlim,
            header_row=['Result', 'Flags', 'Type', 'File'],
            column_width=[10, 5, 4, 56])

    # TODO: handle errors in xglob, to continue processing the next files
    for container, filename, data in xglob.iter_files(args, recursive=options.recursive,
                                                      zip_password=options.zip_password, zip_fname=options.zip_fname):
        # ignore directory names stored in zip files:
        if container and filename.endswith('/'):
            continue
        full_name = '%s in %s' % (filename, container) if container else filename
        # try:
        #     # Open the file
        #     if data is None:
        #         data = open(filename, 'rb').read()
        # except:
        #     log.exception('Error when opening file %r' % full_name)
        #     continue
        if isinstance(data, Exception):
            result = '* ERROR *'
            result_color = 'yellow'
            t.write_row([result, '', '', full_name],
                        colors=[result_color, None, None, None])
            t.write_row(['', '', '', str(data)],
                        colors=[None, None, None, result_color])
        else:
            try:
                vba_parser = olevba.VBA_Parser(filename=filename, data=data, container=container)
                filetype = TYPE2TAG[vba_parser.type]
            except Exception as e:
                # log.error('Error when parsing VBA macros from file %r' % full_name)
                result = '* ERROR *'
                result_color = 'yellow'
                t.write_row([result, '', TYPE2TAG[vba_parser.type], full_name],
                            colors=[result_color, None, None, None])
                t.write_row(['', '', '', str(e)],
                            colors=[None, None, None, result_color])
                continue
            if vba_parser.detect_vba_macros():
                vba_code_all_modules = ''
                try:
                    for (subfilename, stream_path, vba_filename, vba_code) in vba_parser.extract_all_macros():
                        vba_code_all_modules += vba_code + '\n'
                except Exception as e:
                    # log.error('Error when parsing VBA macros from file %r' % full_name)
                    result = '* ERROR *'
                    result_color = 'yellow'
                    t.write_row([result, '', TYPE2TAG[vba_parser.type], full_name],
                                colors=[result_color, None, None, None])
                    t.write_row(['', '', '', str(e)],
                                colors=[None, None, None, result_color])
                    continue
                mraptor = MacroRaptor(vba_code_all_modules)
                mraptor.scan()
                if mraptor.suspicious:
                    result = 'SUSPICIOUS'
                    result_color = 'red'
                else:
                    result = 'Macro OK'
                    result_color = 'cyan'
                t.write_row([result, mraptor.get_flags(), filetype, full_name],
                            colors=[result_color, None, None, None])
                if mraptor.matches and options.show_matches:
                    t.write_row(['', '', '', 'Matches: %r' % mraptor.matches])
            else:
                result = 'No Macro'
                result_color = 'green'
                t.write_row([result, '', filetype, full_name],
                            colors=[result_color, None, None, None])

    print ''
    print 'Flags: A=AutoExec, W=Write, X=Execute'

if __name__ == '__main__':
    main()

# Soundtrack: "Dark Child" by Marlon Williams
