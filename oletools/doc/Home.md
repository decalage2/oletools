python-oletools v0.06 documentation
===================================

This is the home page of the documentation for python-oletools. The latest version can be found [online](https://bitbucket.org/decalage/oletools/wiki), otherwise a copy is provided in the doc subfolder of the package.

[python-oletools](http://www.decalage.info/python/oletools) is a package of python tools to analyze [Microsoft OLE2 files (also called Structured Storage, Compound File Binary Format or Compound Document File Format)](http://en.wikipedia.org/wiki/Compound_File_Binary_Format), such as Microsoft Office documents or Outlook messages, mainly for malware analysis and debugging. It is based on the [OleFileIO_PL](http://www.decalage.info/python/olefileio) parser. See [http://www.decalage.info/python/oletools](http://www.decalage.info/python/oletools) for more info.  

**Quick links:** [Home page](http://www.decalage.info/python/oletools) - [Download](https://bitbucket.org/decalage/oletools/downloads) - [Documentation](https://bitbucket.org/decalage/oletools/wiki) - [Report Issues/Suggestions/Questions](https://bitbucket.org/decalage/oletools/issues?status=new&status=open) - [Contact the author](http://decalage.info/contact) - [Repository](https://bitbucket.org/decalage/oletools) - [Updates on Twitter](https://twitter.com/decalage2)

Note: python-oletools is not related to OLETools published by BeCubed Software.

Tools in python-oletools:
-------------------------

- **[[olebrowse]]**: A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint documents), to
  view and extract individual data streams.
- **[[oleid]]**: a tool to analyze OLE files to detect specific characteristics usually found in malicious files.
- **[[olemeta]]**: a tool to extract all standard properties (metadata) from OLE files.
- **[[oletimes]]**: a tool to extract creation and modification timestamps of all streams and storages.
- **[[olevba]]**: a tool to extract VBA Macro source code from MS Office documents (OLE and OpenXML).
- **[[pyxswf]]**: a tool to detect, extract and analyze Flash objects (SWF) that may
  be embedded in files such as MS Office documents (e.g. Word, Excel) and RTF,
  which is especially useful for malware analysis.
- **[[rtfobj]]**: a tool and python module to extract embedded objects from RTF files.
- and a few others (coming soon)

--------------------------------------------------------------------------

python-oletools documentation
-----------------------------

- [[Home]]
- [[License]]
- [[Install]]
- [[Contribute]], Suggest Improvements or Report Issues
- Tools:
	- [[olebrowse]]
	- [[oleid]]
	- [[olemeta]]
	- [[oletimes]]
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]] 