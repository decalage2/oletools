oletools
========

`oletools <http://www.decalage.info/python/oletools>`_ is a package of
python tools to analyze `Microsoft OLE2 files (also called Structured
Storage, Compound File Binary Format or Compound Document File
Format) <http://en.wikipedia.org/wiki/Compound_File_Binary_Format>`_,
such as Microsoft Office documents or Outlook messages, mainly for
malware analysis and debugging. It is based on the
`OleFileIO\_PL <http://www.decalage.info/python/olefileio>`_ parser. See
`http://www.decalage.info/python/oletools <http://www.decalage.info/python/oletools>`_
for more info.

Tools in oletools:
------------------

-  **olebrowse**: A simple GUI to browse OLE files (e.g. MS Word, Excel,
   Powerpoint documents), to view and extract individual data streams.
-  **pyxswf**: a script to detect, extract and analyze Flash objects
   (SWF) that may be embedded in files such as MS Office documents (e.g.
   Word, Excel), which is especially useful for malware analysis.
-  and a few others (coming soon)

News
----

-  2012-10-09: Initial version of olebrowse and pyxswf
-  see changelog in source code for more info.

Download:
---------

The archive is available on `the project
page <https://bitbucket.org/decalage/oletools/downloads>`_.

olebrowse:
----------

A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint
documents), to view and extract individual data streams.

::

    Usage: olebrowse.py [file]

If you provide a file it will be opened, else a dialog will allow you to
browse folders to open a file. Then if it is a valid OLE file, the list
of data streams will be displayed. You can select a stream, and then
either view its content in a builtin hexadecimal viewer, or save it to a
file for further analysis.

olebrowse project website:
`http://www.decalage.info/python/olebrowse <http://www.decalage.info/python/olebrowse>`_

pyxswf:
-------

pyxswf is a script to detect, extract and analyze Flash objects (SWF
files) that may be embedded in files such as MS Office documents (e.g.
Word, Excel), which is especially useful for malware analysis.

pyxswf is an improved version of xxxswf.py published by Alexander Hanel
on
`http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html <http://hooked-on-mnemonics.blogspot.nl/2011/12/xxxswfpy.html>`_

Compared to xxxswf, it can extract streams from MS Office documents by
parsing their OLE structure properly, which is necessary when streams
are fragmented. Stream fragmentation is a known obfuscation technique,
as explained on
`http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/ <http://www.breakingpointsystems.com/resources/blog/evasion-with-ole2-fragmentation/>`_

For this, simply add the -o option to work on OLE streams rather than
raw files.

::

    Usage: pyxswf.py [options] <file.bad>

    Options:
      -o, --ole             Parse an OLE file (e.g. Word, Excel) to look for SWF
                            in each stream
      -x, --extract         Extracts the embedded SWF(s), names it MD5HASH.swf &
                            saves it in the working dir. No addition args needed
      -h, --help            show this help message and exit
      -y, --yara            Scans the SWF(s) with yara. If the SWF(s) is
                            compressed it will be deflated. No addition args
                            needed
      -s, --md5scan         Scans the SWF(s) for MD5 signatures. Please see func
                            checkMD5 to define hashes. No addition args needed
      -H, --header          Displays the SWFs file header. No addition args needed
      -d, --decompress      Deflates compressed SWFS(s)
      -r PATH, --recdir=PATH
                            Will recursively scan a directory for files that
                            contain SWFs. Must provide path in quotes
      -c, --compress        Compresses the SWF using Zlib

Example - detecting and extracting a SWF file from a Word document on
Windows:

::

    C:\oletools>pyxswf.py -o word_flash.doc
    OLE stream: 'Contents'
    [SUMMARY] 1 SWF(s) in MD5:993664cc86f60d52d671b6610813cfd1:Contents
            [ADDR] SWF 1 at 0x8  - FWS Header

    C:\oletools>pyxswf.py -xo word_flash.doc
    OLE stream: 'Contents'
    [SUMMARY] 1 SWF(s) in MD5:993664cc86f60d52d671b6610813cfd1:Contents
            [ADDR] SWF 1 at 0x8  - FWS Header
                    [FILE] Carved SWF MD5: 2498e9c0701dc0e461ab4358f9102bc5.swf

pyxswf project website:
`http://www.decalage.info/python/pyxswf <http://www.decalage.info/python/pyxswf>`_

How to contribute:
------------------

The code is available in `a Mercurial repository on
bitbucket <https://bitbucket.org/decalage/oletools>`_. You may use it to
submit enhancements or to report any issue.

If you would like to help us improve this module, or simply provide
feedback, you may also send an e-mail to decalage(at)laposte.net.

How to report bugs:
-------------------

To report a bug or any issue, please use the `issue reporting
page <https://bitbucket.org/decalage/olefileio_pl/issues?status=new&status=open>`_,
or send an e-mail with all the information and files to reproduce the
problem.

License
-------

This license applies to the oletools package, apart from the thirdparty
folder which contains third-party files published with their own
license.

The oletools package is copyright (c) 2012, Philippe Lagadec
(http://www.decalage.info) All rights reserved.

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
