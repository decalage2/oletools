olebrowse
=========

olebrowse is a simple GUI to browse OLE files (e.g. MS Word, Excel, Powerpoint documents), to
view and extract individual data streams.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

Usage
-----

	olebrowse.py [file]

If you provide a file it will be opened, else a dialog will allow you to browse folders to open a file. Then if it is a valid OLE file, the list of data streams will be displayed. You can select a stream, and then either view its content in a builtin hexadecimal viewer, or save it to a file for further analysis.

Screenshots
-----------

Main menu, showing all streams in the OLE file:

![](olebrowse1_menu.png)

Menu with actions for a stream:

![](olebrowse2_stream.png)

Hex view for a stream:

![](olebrowse3_hexview.png)

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
