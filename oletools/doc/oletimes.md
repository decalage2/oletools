oletimes
========

oletimes is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract creation and modification times of all streams and storages
in the OLE file.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Usage

	:::text
	oletimes.py <file>

### Example

Checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

	:::text
	>oletimes.py DIAN_caso-5415.doc

	- Root mtime=2014-05-14 12:45:24.752000 ctime=None
	- '\x01CompObj': mtime=None ctime=None
	- '\x05DocumentSummaryInformation': mtime=None ctime=None
	- '\x05SummaryInformation': mtime=None ctime=None
	- '1Table': mtime=None ctime=None
	- 'Data': mtime=None ctime=None
	- 'Macros': mtime=2014-05-14 12:45:24.708000 ctime=2014-05-14 12:45:24.355000
	- 'Macros/PROJECT': mtime=None ctime=None
	- 'Macros/PROJECTwm': mtime=None ctime=None
	- 'Macros/VBA': mtime=2014-05-14 12:45:24.684000 ctime=2014-05-14 12:45:24.355000
	- 'Macros/VBA/ThisDocument': mtime=None ctime=None
	- 'Macros/VBA/_VBA_PROJECT': mtime=None ctime=None
	- 'Macros/VBA/__SRP_0': mtime=None ctime=None
	- 'Macros/VBA/__SRP_1': mtime=None ctime=None
	- 'Macros/VBA/__SRP_2': mtime=None ctime=None
	- 'Macros/VBA/__SRP_3': mtime=None ctime=None
	- 'Macros/VBA/dir': mtime=None ctime=None
	- 'WordDocument': mtime=None ctime=None

## How to use oletimes in Python applications	

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