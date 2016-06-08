python-oletools
===============

`python-oletools <http://www.decalage.info/python/oletools>`__ is a
package of python tools to analyze `Microsoft OLE2
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

Note: python-oletools is not related to OLETools published by BeCubed
Software.

News
----

-  **2016-04-19 v0.46**:
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ does
   not deobfuscate VBA expressions by default (much faster), new option
   --deobf to enable it. Fixed color display bug on Windows for several
   tools.
-  2016-04-12 v0.45: improved
   `rtfobj <https://github.com/decalage2/oletools/wiki/rtfobj>`__ to
   handle several `anti-analysis
   tricks <http://www.decalage.info/rtf_tricks>`__, improved
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ to
   export results in JSON format.
-  2016-03-11 v0.44: improved
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ to
   extract and analyse strings from VBA Forms.
-  2016-03-04 v0.43: added new tool MacroRaptor (mraptor) to detect
   malicious macros, bugfix and slight improvements in
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__.
-  2016-02-07 v0.42: added two new tools oledir and olemap, better
   handling of malformed files and several bugfixes in
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__,
   improved display for
   `olemeta <https://github.com/decalage2/oletools/wiki/olemeta>`__.
-  2015-09-22 v0.41: added new --reveal option to
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__, to
   show the macro code with VBA strings deobfuscated.
-  2015-09-17 v0.40: Improved macro deobfuscation in
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__, to
   decode Hex and Base64 within VBA expressions. Display printable
   deobfuscated strings by default. Improved the VBA\_Parser API.
   Improved performance. Fixed `issue
   #23 <https://github.com/decalage2/oletools/issues/23>`__ with
   sys.stderr.
-  2015-06-19 v0.12:
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ can
   now deobfuscate VBA expressions with any combination of Chr, Asc,
   Val, StrReverse, Environ, +, &, using a VBA parser built with
   `pyparsing <http://pyparsing.wikispaces.com>`__. New options to
   display only the analysis results or only the macros source code. The
   analysis is now done on all the VBA modules at once.
-  2015-05-29 v0.11: Improved parsing of MHTML and ActiveMime/MSO files
   in `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__,
   added several suspicious keywords to VBA scanner (thanks to @ozhermit
   and Davy Douhine for the suggestions)
-  2015-05-06 v0.10:
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ now
   supports Word MHTML files with macros, aka "Single File Web Page"
   (.mht) - see `issue
   #10 <https://github.com/decalage2/oletools/issues/10>`__ for more
   info
-  2015-03-23 v0.09:
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ now
   supports Word 2003 XML files, added anti-sandboxing/VM detection
-  2015-02-08 v0.08:
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ can
   now decode strings obfuscated with Hex/StrReverse/Base64/Dridex and
   extract IOCs. Added new triage mode, support for non-western
   codepages with olefile 0.42, improved API and display, several
   bugfixes.
-  2015-01-05 v0.07: improved
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__ to
   detect suspicious keywords and IOCs in VBA macros, can now scan
   several files and open password-protected zip archives, added a
   Python API, upgraded OleFileIO\_PL to olefile v0.41
-  2014-08-28 v0.06: added
   `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__, a new
   tool to extract VBA Macro source code from MS Office documents
   (97-2003 and 2007+). Improved
   `documentation <https://github.com/decalage2/oletools/wiki>`__
-  2013-07-24 v0.05: added new tools
   `olemeta <https://github.com/decalage2/oletools/wiki/olemeta>`__ and
   `oletimes <https://github.com/decalage2/oletools/wiki/oletimes>`__
-  2013-04-18 v0.04: fixed bug in rtfobj, added documentation for
   `rtfobj <https://github.com/decalage2/oletools/wiki/rtfobj>`__
-  2012-11-09 v0.03: Improved
   `pyxswf <https://github.com/decalage2/oletools/wiki/pyxswf>`__ to
   extract Flash objects from RTF
-  2012-10-29 v0.02: Added
   `oleid <https://github.com/decalage2/oletools/wiki/oleid>`__
-  2012-10-09 v0.01: Initial version of
   `olebrowse <https://github.com/decalage2/oletools/wiki/olebrowse>`__
   and pyxswf
-  see changelog in source code for more info.

Tools in python-oletools:
-------------------------

-  `olebrowse <https://github.com/decalage2/oletools/wiki/olebrowse>`__:
   A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint
   documents), to view and extract individual data streams.
-  `oleid <https://github.com/decalage2/oletools/wiki/oleid>`__: to
   analyze OLE files to detect specific characteristics usually found in
   malicious files.
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
-  `olevba <https://github.com/decalage2/oletools/wiki/olevba>`__: to
   extract and analyze VBA Macro source code from MS Office documents
   (OLE and OpenXML).
-  `MacroRaptor <https://github.com/decalage2/oletools/wiki/mraptor>`__:
   to detect malicious VBA Macros
-  `pyxswf <https://github.com/decalage2/oletools/wiki/pyxswf>`__: to
   detect, extract and analyze Flash objects (SWF) that may be embedded
   in files such as MS Office documents (e.g. Word, Excel) and RTF,
   which is especially useful for malware analysis.
-  `oleobj <https://github.com/decalage2/oletools/wiki/oleobj>`__: to
   extract embedded objects from OLE files.
-  `rtfobj <https://github.com/decalage2/oletools/wiki/rtfobj>`__: to
   extract embedded objects from RTF files.
-  and a few others (coming soon)

Download and Install:
---------------------

To use python-oletools from the command line as analysis tools, you may
simply `download the latest release
archive <https://github.com/decalage2/oletools/releases>`__ and extract
the files into the directory of your choice.

You may also download the `latest development
version <https://github.com/decalage2/oletools/archive/master.zip>`__
with the most recent features.

Another possibility is to use a git client to clone the repository
(https://github.com/decalage2/oletools.git) into a folder. You can then
update it easily in the future.

If you plan to use python-oletools with other Python applications or
your own scripts, then the simplest solution is to use "**pip install
oletools**\ " or "**easy\_install oletools**\ " to download and install
in one go. Otherwise you may download/extract the zip archive and run
"**setup.py install**\ ".

**Important: to update oletools** if it is already installed, you must
run **"pip install -U oletools"**, otherwise pip will not update it.

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

The python-oletools package is copyright (c) 2012-2016 Philippe Lagadec
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
