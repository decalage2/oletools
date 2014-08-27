oleid
=====

oleid is a script to analyze OLE files such as MS Office documents (e.g. Word,
Excel), to detect specific characteristics usually found in malicious files (e.g. malware).
For example it can detect VBA macros and embedded Flash objects.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Main Features

- Detect OLE file type from its internal structure (e.g. MS Word, Excel, PowerPoint, ...)
- Detect VBA Macros
- Detect embedded Flash objects
- Detect embedded OLE objects
- Detect MS Office encryption
- Can be used as a command-line tool
- Python API to integrate it in your applications

Planned improvements:

- Extract the most important metadata fields
- Support for OpenXML files and embedded OLE files
- Generic VBA macros detection
- Detect auto-executable VBA macros
- Extended OLE file types detection
- Detect unusual OLE structures (fragmentation, unused sectors, etc)
- Options to scan multiple files
- Options to scan files from encrypted zip archives
- CSV output

## Usage

	:::text
	oleid.py <file>

### Example 

Analyzing a Word document containing a Flash object and VBA macros:

	:::text
	C:\oletools>oleid.py word_flash_vba.doc

	Filename: word_flash_vba.doc
	OLE format: True
	Has SummaryInformation stream: True
	Application name: Microsoft Office Word
	Encrypted: False
	Word Document: True
	VBA Macros: True
	Excel Workbook: False
	PowerPoint Presentation: False
	Visio Drawing: False
	ObjectPool: True
	Flash objects: 1

## How to use oleid in Python applications	

TODO

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