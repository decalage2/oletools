"""
clsid.py

Collection of known CLSIDs and related vulnerabilities, for the oletools.

Author: Philippe Lagadec - http://www.decalage.info
License: BSD, see source code or documentation

clsid is part of the python-oletools package:
http://www.decalage.info/python/oletools
"""

#=== LICENSE ==================================================================

# oletools are copyright (c) 2018 Philippe Lagadec (http://www.decalage.info)
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
# 2018-04-11 v0.53 PL: - added collection of CLSIDs
# 2018-04-13       PL: - moved KNOWN_CLSIDS from oledir to common.clsid
#                  SQ: - several additions by Shiao Qu
# 2018-04-18       PL: - added known-bad CLSIDs from Cuckoo sandbox (issue #290)

__version__ = '0.53dev5'


KNOWN_CLSIDS = {
    # MS Office files
    '00020906-0000-0000-C000-000000000046': 'Microsoft Word 97-2003 Document',
    '00020900-0000-0000-C000-000000000046': 'Microsoft Word 6.0-7.0 Document',
    '00020832-0000-0000-C000-000000000046': 'Excel sheet with macro enabled',
    '00020833-0000-0000-C000-000000000046': 'Excel binary sheet with macro enabled',

    # OLE Objects
    '00000300-0000-0000-C000-000000000046': 'StdOleLink (embedded OLE object)',
    'D7053240-CE69-11CD-A777-00DD01143C57': 'Microsoft Forms 2.0 CommandButton',
    '0002CE02-0000-0000-C000-000000000046': 'MS Equation Editor (may trigger CVE-2017-11882 or CVE-2018-0802)',
    'F20DA720-C02F-11CE-927B-0800095AE340': 'Package (may contain and run any file)',
    '0003000C-0000-0000-C000-000000000046': 'Package (may contain and run any file)',
    'D27CDB6E-AE6D-11CF-96B8-444553540000': 'Shockwave Flash Object (may trigger many CVEs)',
    # Known-bad CLSIDs from Cuckoo Sandbox:
    # https://github.com/cuckoosandbox/community/blob/master/modules/signatures/windows/office.py#L314
    "BDD1F04B-858B-11D1-B16A-00C0F0283628": "MSCOMCTL.ListViewCtrl (may trigger CVE-2012-0158)",
    "996BF5E0-8044-4650-ADEB-0B013914E99C": "MSCOMCTL.ListViewCtrl (may trigger CVE-2012-0158)",
    "C74190B6-8589-11d1-B16A-00C0F0283628": "MSCOMCTL.TreeCtrl (may trigger CVE-2012-0158)",
    "9181DC5F-E07D-418A-ACA6-8EEA1ECB8E9E": "MSCOMCTL.TreeCtrl (may trigger CVE-2012-0158)",
    "1EFB6596-857C-11D1-B16A-00C0F0283628": "MSCOMCTL.TabStrip (may trigger CVE-2012-1856, CVE-2013-3906)",
    "66833FE6-8583-11D1-B16A-00C0F0283628": "MSCOMCTL.Toolbar (may trigger CVE-2012-1856)",
    "DD9DA666-8594-11D1-B16A-00C0F0283628": "MSCOMCTL.ImageComboCtrl (may trigger CVE-2014-1761)",
    "00000535-0000-0010-8000-00AA006D2EA4": "ADODB.RecordSet (may trigger CVE-2015-0097)",
    "0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC": "MSScriptControl.ScriptControl (may trigger CVE-2015-0097)",
    "05741520-C4EB-440A-AC3F-9643BBC9F847": "otkloadr.WRLoader (may trigger CVE-2015-1641)",
    'A08A033D-1A75-4AB6-A166-EAD02F547959': 'otkloadr CWRAssembly Object (may trigger CVE-2015-1641)',
    # TODO "F4754C9B-64F5-4B40-8AF4-679732AC0607": "Microsoft Word Document (may trigger CVE-2015-1641)", ???
    "4C599241-6926-101B-9992-00000B65C6F9": "Forms.Image (may trigger CVE-2015-2424)",
    "44F9A03B-A3EC-4F3B-9364-08E0007F21DF": "Control.TaskSymbol (may trigger CVE-2015-2424)",
    '3050F4D8-98B5-11CF-BB82-00AA00BDCE0B': 'HTML Application (may trigger CVE-2017-0199)',

    # Monikers
    '00000303-0000-0000-C000-000000000046': 'File Moniker (may trigger CVE-2017-0199 or CVE-2017-8570)',
    '00000304-0000-0000-C000-000000000046': 'Item Moniker',
    '00000305-0000-0000-C000-000000000046': 'Anti Moniker',
    '00000306-0000-0000-C000-000000000046': 'Pointer Moniker',
    '00000308-0000-0000-C000-000000000046': 'Packager Moniker',
    '00000309-0000-0000-C000-000000000046': 'Composite Moniker (may trigger CVE-2017-8570)',
    '0000031a-0000-0000-C000-000000000046': 'Class Moniker',
    '0002034c-0000-0000-C000-000000000046': 'OutlookAttachMoniker',
    '0002034e-0000-0000-C000-000000000046': 'OutlookMessageMoniker',
    '79EAC9E0-BAF9-11CE-8C82-00AA004BA90B': 'URL Moniker (may trigger CVE-2017-0199 or CVE-2017-8570)',
    'ECABB0C7-7F19-11D2-978E-0000F8757E2A': 'SOAP Moniker (may trigger CVE-2017-8759)',
    'ECABAFC6-7F19-11D2-978E-0000F8757E2A': 'New Moniker',
    # ref: https://justhaifei1.blogspot.nl/2017/07/bypassing-microsofts-cve-2017-0199-patch.html
    '06290BD2-48AA-11D2-8432-006008C3FBFC': 'Factory bindable using IPersistMoniker (scripletfile)',
    '06290BD3-48AA-11D2-8432-006008C3FBFC': 'Script Moniker, aka Moniker to a Windows Script Component (may trigger CVE-2017-0199)',
}

