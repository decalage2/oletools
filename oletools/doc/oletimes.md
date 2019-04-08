oletimes
========

oletimes is a script to parse OLE files such as MS Office documents (e.g. Word,
Excel), to extract creation and modification times of all streams and storages
in the OLE file.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Usage

```text
oletimes <file>
```

### Example

Checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

```text
>oletimes DIAN_caso-5415.doc

+----------------------------+---------------------+---------------------+
| Stream/Storage name        | Modification Time   | Creation Time       |
+----------------------------+---------------------+---------------------+
| Root                       | 2014-05-14 12:45:24 | None                |
| '\x01CompObj'              | None                | None                |
| '\x05DocumentSummaryInform | None                | None                |
| ation'                     |                     |                     |
| '\x05SummaryInformation'   | None                | None                |
| '1Table'                   | None                | None                |
| 'Data'                     | None                | None                |
| 'Macros'                   | 2014-05-14 12:45:24 | 2014-05-14 12:45:24 |
| 'Macros/PROJECT'           | None                | None                |
| 'Macros/PROJECTwm'         | None                | None                |
| 'Macros/VBA'               | 2014-05-14 12:45:24 | 2014-05-14 12:45:24 |
| 'Macros/VBA/ThisDocument'  | None                | None                |
| 'Macros/VBA/_VBA_PROJECT'  | None                | None                |
| 'Macros/VBA/__SRP_0'       | None                | None                |
| 'Macros/VBA/__SRP_1'       | None                | None                |
| 'Macros/VBA/__SRP_2'       | None                | None                |
| 'Macros/VBA/__SRP_3'       | None                | None                |
| 'Macros/VBA/dir'           | None                | None                |
| 'WordDocument'             | None                | None                |
+----------------------------+---------------------+---------------------+
```

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
