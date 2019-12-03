python-oletools
===============

|PyPI| |Build Status| |Say Thanks!|

`oletools <http://www.decalage.info/python/oletools>`__ is a package of
python tools to analyze `Microsoft OLE2
files <http://en.wikipedia.org/wiki/Compound_File_Binary_Format>`__
(also called Structured Storage, Compound File Binary Format or Compound
Document File Format), such as Microsoft Office documents or Outlook
messages, mainly for malware analysis, forensics and debugging. It is
based on the `olefile <http://www.decalage.info/olefile>`__ parser. See
http://www.decalage.info/python/oletools for more info.

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

-  **2019-12-03 v0.55**:

   -  olevba:

      -  added support for SLK files and XLM macro extraction from SLK
      -  VBA Stomping detection
      -  integrated pcodedmp to extract and disassemble P-code
      -  detection of suspicious keywords and IOCs in P-code
      -  new option --pcode to display P-code disassembly
      -  improved detection of auto execution triggers

   -  rtfobj: added URL carver for CVE-2017-0199
   -  better handling of unicode for systems with locale that does not
      support UTF-8, e.g. LANG=C (PR #365)
   -  tests:

      -  test files can now be encrypted, to avoid antivirus alerts (PR
         #217, issue #215)
      -  tests that trigger antivirus alerts have been temporarily
         disabled (issue #215)

-  **2019-05-22 v0.54.2**:

   -  bugfix release: fixed several issues related to encrypted
      documents and XLM/XLF Excel 4 macros
   -  msoffcrypto-tool is now installed by default to handle encrypted
      documents
   -  olevba and msodde now handle documents encrypted with common
      passwords such as 123, 1234, 4321, 12345, 123456, VelvetSweatShop
      automatically.

-  **2019-04-04 v0.54**:

   -  olevba, msodde: added support for encrypted MS Office files
   -  olevba: added detection and extraction of XLM/XLF Excel 4 macros
      (thanks to plugin_biff from Didier Stevens' oledump)
   -  olevba, mraptor: added detection of VBA running Excel 4 macros
   -  olevba: detect and display special characters such as backspace
   -  olevba: colorized output showing suspicious keywords in the VBA
      code
   -  olevba, mraptor: full Python 3 compatibility, no separate
      olevba3/mraptor3 anymore
   -  olevba: improved handling of code pages and unicode
   -  olevba: fixed a false-positive in VBA macro detection
   -  rtfobj: improved OLE Package handling, improved Equation object
      detection
   -  oleobj: added detection of external links to objects in OpenXML
   -  replaced third party packages by PyPI dependencies

-  2018-05-30 v0.53:

   -  olevba and mraptor can now parse Word/PowerPoint 2007+ pure XML
      files (aka Flat OPC format)
   -  improved support for VBA forms in olevba (oleform)
   -  rtfobj now displays the CLSID of OLE objects, which is the best
      way to identify them. Known-bad CLSIDs such as MS Equation Editor
      are highlighted in red.
   -  Updated rtfobj to handle obfuscated RTF samples.
   -  rtfobj now handles the "\'" obfuscation trick seen in recent
      samples such as
      https://twitter.com/buffaloverflow/status/989798880295444480, by
      emulating the MS Word bug described in
      https://securelist.com/disappearing-bytes/84017/
   -  msodde: improved detection of DDE formulas in CSV files
   -  oledir now displays the tree of storage/streams, along with CLSIDs
      and their meaning.
   -  common.clsid contains the list of known CLSIDs, and their links to
      CVE vulnerabilities when relevant.
   -  oleid now detects encrypted OpenXML files
   -  fixed bugs in oleobj, rtfobj, oleid, olevba

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
`Anlyz.io <https://sandbox.anlyz.io/>`__,
`AssemblyLine <https://www.cse-cst.gc.ca/en/assemblyline>`__,
`CAPE <https://github.com/ctxis/CAPE>`__, `Cuckoo
Sandbox <https://github.com/cuckoosandbox/cuckoo>`__,
`DARKSURGEON <https://github.com/cryps1s/DARKSURGEON>`__,
`Deepviz <https://sandbox.deepviz.com/>`__,
`dridex.malwareconfig.com <https://dridex.malwareconfig.com>`__,
`FAME <https://certsocietegenerale.github.io/fame/>`__,
`FLARE-VM <https://github.com/fireeye/flare-vm>`__,
`Hybrid-analysis.com <https://www.hybrid-analysis.com/>`__, `Joe
Sandbox <https://www.document-analyzer.net/>`__, `Laika
BOSS <https://github.com/lmco/laikaboss>`__,
`MacroMilter <https://github.com/sbidy/MacroMilter>`__,
`mailcow <https://mailcow.email/>`__,
`malshare.io <https://malshare.io>`__,
`malware-repo <https://github.com/Tigzy/malware-repo>`__, `Malware
Repository Framework (MRF) <https://www.adlice.com/download/mrf/>`__,
`olefy <https://github.com/HeinleinSupport/olefy>`__,
`PeekabooAV <https://github.com/scVENUS/PeekabooAV>`__,
`pcodedmp <https://github.com/bontchev/pcodedmp>`__,
`PyCIRCLean <https://github.com/CIRCL/PyCIRCLean>`__,
`REMnux <https://remnux.org/>`__,
`Snake <https://github.com/countercept/snake>`__,
`SNDBOX <https://app.sndbox.com>`__,
`Strelka <https://github.com/target/strelka>`__,
`stoQ <https://stoq.punchcyber.com/>`__,
`TheHive/Cortex <https://github.com/TheHive-Project/Cortex-Analyzers>`__,
`TSUGURI Linux <https://tsurugi-linux.org/>`__,
`Vba2Graph <https://github.com/MalwareCantFly/Vba2Graph>`__,
`Viper <http://viper.li/>`__,
`ViperMonkey <https://github.com/decalage2/ViperMonkey>`__,
`YOMI <https://yomi.yoroi.company>`__, and probably
`VirusTotal <https://www.virustotal.com>`__. And quite a few `other
projects on
GitHub <https://github.com/search?q=oletools&type=Repositories>`__.
(Please `contact me <(http://decalage.info/contact)>`__ if you have or
know a project using oletools)

Download and Install:
---------------------

The recommended way to download and install/update the **latest stable
release** of oletools is to use
`pip <https://pip.pypa.io/en/stable/installing/>`__:

-  On Linux/Mac: ``sudo -H pip install -U oletools``
-  On Windows: ``pip install -U oletools``

This should automatically create command-line scripts to run each tool
from any directory: ``olevba``, ``mraptor``, ``rtfobj``, etc.

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

The python-oletools package is copyright (c) 2012-2019 Philippe Lagadec
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
