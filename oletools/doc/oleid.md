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

```text
oleid <file>
```

### Example 

Analyzing a Word document containing a Flash object and VBA macros:

```text
C:\oletools>oleid word_flash_vba.doc

Filename: word_flash_vba.doc
+-------------------------------+-----------------------+
| Indicator                     | Value                 |
+-------------------------------+-----------------------+
| OLE format                    | True                  |
| Has SummaryInformation stream | True                  |
| Application name              | Microsoft Office Word |
| Encrypted                     | False                 |
| Word Document                 | True                  |
| VBA Macros                    | True                  |
| Excel Workbook                | False                 |
| PowerPoint Presentation       | False                 |
| Visio Drawing                 | False                 |
| ObjectPool                    | True                  |
| Flash objects                 | 1                     |
+-------------------------------+-----------------------+
```

## How to use oleid in your Python applications	

First, import oletools.oleid, and create an **OleID** object to scan a file:

```python
import oletools.oleid

oid = oletools.oleid.OleID(filename)
```

Note: filename can be a filename, a file-like object, or a bytes string containing the file to be analyzed.

Second, call the **check()** method. It returns a list of **Indicator** objects.

Each Indicator object has the following attributes:

- **id**: str, identifier for the indicator
- **name**: str, name to display the indicator
- **description**: str, long description of the indicator
- **type**: class of the indicator (e.g. bool, str, int)
- **value**: value of the indicator

For example, the following code displays all the indicators:

```python
indicators = oid.check()
for i in indicators:
    print 'Indicator id=%s name="%s" type=%s value=%s' % (i.id, i.name, i.type, repr(i.value))
    print 'description:', i.description
    print ''
```

See the source code of oleid.py for more details.

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
