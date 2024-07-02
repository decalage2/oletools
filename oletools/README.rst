python-oletools
===============

|PyPI| |Build Status| |Say Thanks!|

`oletools <http://www.decalage.info/python/oletools>`__ is a package of
python tools to analyze `Microsoft OLE2
files <http://en.wikipedia.org/wiki/Compound_File_Binary_Format>`__
(also called Structured Storage, Compound File Binary Format or Compound
Document File Format), such as Microsoft Office 97-2003 documents, MSI
files or Outlook messages, mainly for malware analysis, forensics and
debugging. It is based on the
`olefile <http://www.decalage.info/olefile>`__ parser.

It also provides tools to analyze RTF files and files based on the
`OpenXML format <https://en.wikipedia.org/wiki/Office_Open_XML>`__ (aka
OOXML) such as MS Office 2007+ documents, XPS or MSIX files.

For example, oletools can detect, extract and analyse VBA macros, OLE
objects, Excel 4 macros (XLM) and DDE links.

See http://www.decalage.info/python/oletools for more info.

**Quick links:** `Home
page <http://www.decalage.info/python/oletools>`__ -
`Download/Install <https://github.com/decalage2/oletools/wiki/Install>`__
- `Documentation <https://github.com/decalage2/oletools/wiki>`__ -
`Report
Issues/Suggestions/Questions <https://github.com/decalage2/oletools/issues>`__
- `Contact the Author <http://decalage.info/contact>`__ -
`Repository <https://github.com/decalage2/oletools>`__ - `Updates on
Twitter <https://twitter.com/decalage2>`__
`Cheatsheet <https://github.com/decalage2/oletools/blob/master/cheatsheet/oletools_cheatsheet.pdf>`__

Note: python-oletools is not related to OLETools published by BeCubed
Software.

News
----

-  **2024-07-02 v0.60.2**:

   -  olevba:

      -  fixed a bug in open_slk (issue #797, PR #769)
      -  fixed a bug due to new PROJECTCOMPATVERSION record in dir
         stream (PR #723, issues #700, #701, #725, #791, #808, #811,
         #833)

   -  oleobj: fixed SyntaxError with Python 3.12 (PR #855),
      SyntaxWarning (PR #774)
   -  rtfobj: fixed SyntaxError with Python 3.12 (PR #854)
   -  clsid: added CLSIDs for MSI, Zed
   -  ftguess: added MSI, PNG and OneNote formats
   -  pyxswf: fixed python 3.12 compatibility (PR #841, issue #813)
   -  setup/requirements: allow pyparsing 3 to solve install issues (PR
      #812, issue #762)

-  **2022-05-09 v0.60.1**:

   -  olevba:

      -  fixed a bug when calling XLMMacroDeobfuscator (PR #737)
      -  removed keyword "sample" causing false positives

   -  oleid: fixed OleID init issue (issue #695, PR #696)
   -  oleobj:

      -  added simple detection of CVE-2021-40444 initial stage
      -  added detection for customUI onLoad
      -  improved handling of incorrect filenames in OLE package (PR
         #451)

   -  rtfobj: fixed code to find URLs in OLE2Link objects for Py3 (issue
      #692)
   -  ftguess:

      -  added PowerPoint and XPS formats (PR #716)
      -  fixed issue with XPS and malformed documents (issue #711)
      -  added XLSB format (issue #758)

   -  improved logging with common module log_helper (PR #449)

-  **2021-06-02 v0.60**:

   -  ftguess: new tool to identify file formats and containers (issue
      #680)
   -  oleid: (issue #679)

      -  each indicator now has a risk level
      -  calls ftguess to identify file formats
      -  calls olevba+mraptor to detect and analyse VBA+XLM macros

   -  olevba:

      -  when XLMMacroDeobfuscator is available, use it to extract and
         deobfuscate XLM macros

   -  rtfobj:

      -  use ftguess to identify file type of OLE Package (issue #682)
      -  fixed bug in re_executable_extensions

   -  crypto: added PowerPoint transparent password '/01Hannes
      Ruescher/01' (issue #627)
   -  setup: XLMMacroDeobfuscator, xlrd2 and pyxlsb2 added as optional
      dependencies

See the `full
changelog <https://github.com/decalage2/oletools/wiki/Changelog>`__ for
more information.

Tools:
------

Tools to analyze malicious documents
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

-  `oleid <https://github.com/decalage2/oletools/wiki/oleid>`__: to
   analyze OLE files to detect specific characteristics usually found in
   malicious files.
-  `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__: to
   extract and analyze VBA Macro source code from MS Office documents
   (OLE and OpenXML).
-  `MacroRaptor <https://github.com/decalage2/oletools/wiki/mraptor>`__:
   to detect malicious VBA Macros
-  `msodde <https://github.com/decalage2/oletools/wiki/msodde>`__: to
   detect and extract DDE/DDEAUTO links from MS Office documents, RTF
   and CSV
-  `pyxswf <https://github.com/decalage2/oletools/wiki/pyxswf>`__: to
   detect, extract and analyze Flash objects (SWF) that may be embedded
   in files such as MS Office documents (e.g. Word, Excel) and RTF,
   which is especially useful for malware analysis.
-  `oleobj <https://github.com/decalage2/oletools/wiki/oleobj>`__: to
   extract embedded objects from OLE files.
-  `rtfobj <https://github.com/decalage2/oletools/wiki/rtfobj>`__: to
   extract embedded objects from RTF files.

Tools to analyze the structure of OLE files
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

-  `olebrowse <https://github.com/decalage2/oletools/wiki/olebrowse>`__:
   A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint
   documents), to view and extract individual data streams.
-  `olemeta <https://github.com/decalage2/oletools/wiki/olemeta>`__: to
   extract all standard properties (metadata) from OLE files.
-  `oletimes <https://github.com/decalage2/oletools/wiki/oletimes>`__:
   to extract creation and modification timestamps of all streams and
   storages.
-  `oledir <https://github.com/decalage2/oletools/wiki/oledir>`__: to
   display all the directory entries of an OLE file, including free and
   orphaned entries.
-  `olemap <https://github.com/decalage2/oletools/wiki/olemap>`__: to
   display a map of all the sectors in an OLE file.

Projects using oletools:
------------------------

oletools are used by a number of projects and online malware analysis
services, including `ACE <https://github.com/IntegralDefense/ACE>`__,
`ADAPT <https://www.blackhat.com/eu-23/briefings/schedule/index.html#unmasking-apts-an-automated-approach-for-real-world-threat-attribution-35162>`__,
`Anlyz.io <https://sandbox.anlyz.io/>`__,
`AssemblyLine <https://www.cse-cst.gc.ca/en/assemblyline>`__, `Binary
Refinery <https://github.com/binref/refinery>`__,
`CAPE <https://github.com/ctxis/CAPE>`__,
`CinCan <https://cincan.io>`__, `Cortex XSOAR (Palo
Alto) <https://cortex.marketplace.pan.dev/marketplace/details/Oletools/>`__,
`Cuckoo Sandbox <https://github.com/cuckoosandbox/cuckoo>`__,
`DARKSURGEON <https://github.com/cryps1s/DARKSURGEON>`__,
`Deepviz <https://sandbox.deepviz.com/>`__,
`DIARIO <https://diario.elevenpaths.com/>`__,
`dridex.malwareconfig.com <https://dridex.malwareconfig.com>`__, `EML
Analyzer <https://github.com/ninoseki/eml_analyzer>`__,
`EXPMON <https://pub.expmon.com/>`__,
`FAME <https://certsocietegenerale.github.io/fame/>`__,
`FLARE-VM <https://github.com/fireeye/flare-vm>`__, `GLIMPS
Malware <https://www.glimps.fr/en/glimps-malware-2/>`__,
`Hybrid-analysis.com <https://www.hybrid-analysis.com/>`__, `InQuest
Labs <https://labs.inquest.net/>`__,
`IntelOwl <https://github.com/certego/IntelOwl>`__, `Joe
Sandbox <https://www.document-analyzer.net/>`__, `Laika
BOSS <https://github.com/lmco/laikaboss>`__,
`MacroMilter <https://github.com/sbidy/MacroMilter>`__,
`mailcow <https://mailcow.email/>`__,
`malshare.io <https://malshare.io>`__,
`malware-repo <https://github.com/Tigzy/malware-repo>`__, `Malware
Repository Framework (MRF) <https://www.adlice.com/download/mrf/>`__,
`MalwareBazaar <https://bazaar.abuse.ch/>`__,
`olefy <https://github.com/HeinleinSupport/olefy>`__,
`Pandora <https://github.com/pandora-analysis/pandora>`__,
`PeekabooAV <https://github.com/scVENUS/PeekabooAV>`__,
`pcodedmp <https://github.com/bontchev/pcodedmp>`__,
`PyCIRCLean <https://github.com/CIRCL/PyCIRCLean>`__,
`QFlow <https://www.quarkslab.com/products-qflow/>`__,
`Qu1cksc0pe <https://github.com/CYB3RMX/Qu1cksc0pe>`__, `Tylabs
QuickSand <https://github.com/tylabs/quicksand>`__,
`REMnux <https://remnux.org/>`__,
`Snake <https://github.com/countercept/snake>`__,
`SNDBOX <https://app.sndbox.com>`__, `Splunk add-on for MS O365
Email <https://splunkbase.splunk.com/app/5365/>`__,
`SpuriousEmu <https://github.com/ldbo/SpuriousEmu>`__,
`Strelka <https://github.com/target/strelka>`__,
`stoQ <https://stoq.punchcyber.com/>`__, `Sublime
Platform/MQL <https://docs.sublimesecurity.com/docs/enrichment-functions>`__,
`Subparse <https://github.com/jstrosch/subparse>`__,
`TheHive/Cortex <https://github.com/TheHive-Project/Cortex-Analyzers>`__,
`ThreatBoook <https://s.threatbook.com/>`__, `TSUGURI
Linux <https://tsurugi-linux.org/>`__,
`Vba2Graph <https://github.com/MalwareCantFly/Vba2Graph>`__,
`Viper <http://viper.li/>`__,
`ViperMonkey <https://github.com/decalage2/ViperMonkey>`__,
`YOMI <https://yomi.yoroi.company>`__, and probably
`VirusTotal <https://www.virustotal.com>`__,
`FileScan.IO <https://www.filescan.io>`__. And quite a few `other
projects on
GitHub <https://github.com/search?q=oletools&type=Repositories>`__.
(Please `contact me <(http://decalage.info/contact)>`__ if you have or
know a project using oletools)

Download and Install:
---------------------

The recommended way to download and install/update the **latest stable
release** of oletools is to use
`pip <https://pip.pypa.io/en/stable/installing/>`__:

-  On Linux/Mac: ``sudo -H pip install -U oletools[full]``
-  On Windows: ``pip install -U oletools[full]``

This should automatically create command-line scripts to run each tool
from any directory: ``olevba``, ``mraptor``, ``rtfobj``, etc.

The keyword ``[full]`` means that all optional dependencies will be
installed, such as XLMMacroDeobfuscator. If you prefer a lighter version
without optional dependencies, just remove ``[full]`` from the command
line.

To get the **latest development version** instead:

-  On Linux/Mac:
   ``sudo -H pip install -U https://github.com/decalage2/oletools/archive/master.zip``
-  On Windows:
   ``pip install -U https://github.com/decalage2/oletools/archive/master.zip``

See the
`documentation <https://github.com/decalage2/oletools/wiki/Install>`__
for other installation options.

Documentation:
--------------

The latest version of the documentation can be found
`online <https://github.com/decalage2/oletools/wiki>`__, otherwise a
copy is provided in the doc subfolder of the package.

How to Suggest Improvements, Report Issues or Contribute:
---------------------------------------------------------

This is a personal open-source project, developed on my spare time. Any
contribution, suggestion, feedback or bug report is welcome.

To suggest improvements, report a bug or any issue, please use the
`issue reporting page <https://github.com/decalage2/oletools/issues>`__,
providing all the information and files to reproduce the problem.

You may also `contact the author <http://decalage.info/contact>`__
directly to provide feedback.

The code is available in `a GitHub
repository <https://github.com/decalage2/oletools>`__. You may use it to
submit enhancements using forks and pull requests.

License
-------

This license applies to the python-oletools package, apart from the
thirdparty folder which contains third-party files published with their
own license.

The python-oletools package is copyright (c) 2012-2024 Philippe Lagadec
(http://www.decalage.info)

All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

-  Redistributions of source code must retain the above copyright
   notice, this list of conditions and the following disclaimer.
-  Redistributions in binary form must reproduce the above copyright
   notice, this list of conditions and the following disclaimer in the
   documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS
IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED
TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A
PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED
TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

--------------

olevba contains modified source code from the officeparser project,
published under the following MIT License (MIT):

officeparser is copyright (c) 2014 John William Davison

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.. |PyPI| image:: https://img.shields.io/pypi/v/oletools.svg
   :target: https://pypi.org/project/oletools/
.. |Build Status| image:: https://travis-ci.org/decalage2/oletools.svg?branch=master
   :target: https://travis-ci.org/decalage2/oletools
.. |Say Thanks!| image:: https://img.shields.io/badge/Say%20Thanks-!-1EAEDB.svg
   :target: https://saythanks.io/to/decalage2
