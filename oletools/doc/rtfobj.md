rtfobj
======

rtfobj is a Python module to extract embedded objects from RTF files, such as
OLE ojects. It can be used as a Python library or a command-line tool.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Usage

```text
rtfobj.py <file.rtf>
```

It extracts and decodes all the data blocks encoded as hexadecimal in the RTF document,
and saves them as files named "object_xxxx.bin", xxxx being the location of the object
in the RTF file.



## How to use rtfobj in Python applications	

Usage as a python module: 

rtf_iter_objects(filename) is an iterator which yields a tuple (index, object) providing the index of each hexadecimal stream in the RTF file, and the corresponding decoded object. 

Example:

```python
import rtfobj
for index, data in rtfobj.rtf_iter_objects("myfile.rtf"):
    print 'found object size %d at index %08X' % (len(data), index)
```

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
	- [[oledir]]
	- [[olemap]]
	- [[olevba]]
	- [[mraptor]]
	- [[pyxswf]]
	- [[oleobj]]
	- [[rtfobj]]
