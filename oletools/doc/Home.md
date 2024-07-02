python-oletools documentation
=============================

This is the home page of the documentation for python-oletools. The latest version can be found 
[online](https://github.com/decalage2/oletools/wiki), otherwise a copy is provided in the doc subfolder of the package.

[oletools](http://www.decalage.info/python/oletools) is a package of python tools to analyze
[Microsoft OLE2 files](http://en.wikipedia.org/wiki/Compound_File_Binary_Format) 
(also called Structured Storage, Compound File Binary Format or Compound Document File Format), 
such as Microsoft Office 97-2003 documents, MSI files or Outlook messages, mainly for malware analysis, 
forensics and debugging. 
It is based on the [olefile](http://www.decalage.info/olefile) parser. 

It also provides tools to analyze RTF files and files based on the [OpenXML format](https://en.wikipedia.org/wiki/Office_Open_XML) 
(aka OOXML) such as MS Office 2007+ documents, XPS or MSIX files.

For example, oletools can detect, extract and analyse VBA macros, OLE objects, Excel 4 macros (XLM) and DDE links.

See [http://www.decalage.info/python/oletools](http://www.decalage.info/python/oletools) for more info.

**Quick links:**
[Home page](http://www.decalage.info/python/oletools) -
[Download/Install](https://github.com/decalage2/oletools/wiki/Install) -
[Documentation](https://github.com/decalage2/oletools/wiki) -
[Report Issues/Suggestions/Questions](https://github.com/decalage2/oletools/issues) -
[Contact the Author](http://decalage.info/contact) -
[Repository](https://github.com/decalage2/oletools) -
[Updates on Twitter](https://twitter.com/decalage2)

Note: python-oletools is not related to OLETools published by BeCubed Software.

Tools in python-oletools:
-------------------------

### Tools to analyze malicious documents

- **[[oleid]]**: to analyze OLE files to detect specific characteristics usually found in malicious files.
- **[[olevba]]**: to extract and analyze VBA Macro source code from MS Office documents (OLE and OpenXML).
- **[[mraptor]]**: to detect malicious VBA Macros
- **[[msodde]]**: to detect and extract DDE/DDEAUTO links from MS Office documents, RTF and CSV
- **[[pyxswf]]**: to detect, extract and analyze Flash objects (SWF) that may
  be embedded in files such as MS Office documents (e.g. Word, Excel) and RTF,
  which is especially useful for malware analysis.
- **[[oleobj]]**: to extract embedded objects from OLE files.
- **[[rtfobj]]**: to extract embedded objects from RTF files.

### Tools to analyze the structure of OLE files 

- **[[olebrowse]]**: A simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint documents), to
  view and extract individual data streams.
- **[[olemeta]]**: to extract all standard properties (metadata) from OLE files.
- **[[oletimes]]**: to extract creation and modification timestamps of all streams and storages.
- **[[oledir]]**: to display all the directory entries of an OLE file, including free and orphaned entries.
- **[[olemap]]**: to display a map of all the sectors in an OLE file.
- and a few others (coming soon)

--------------------------------------------------------------------------

python-oletools documentation
-----------------------------

- [[Home]]
- [[License]]
- [[Install]]
- [[Contribute]], Suggest Improvements or Report Issues
- Tools:
	- [[mraptor]]
	- [[msodde]]
	- [[olebrowse]]
	- [[oledir]]
	- [[oleid]]
	- [[olemap]]
	- [[olemeta]]
	- [[oleobj]]
	- [[oletimes]]
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]]
