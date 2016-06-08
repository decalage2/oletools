How to Download and Install python-oletools
===========================================

Pre-requisites
--------------

For now, python-oletools require **Python 2.x**, if possible 2.7 or 2.6 to enable all features. 

They are not compatible with Python 3.x yet. (Please contact me if that is a strong requirement)


To use oletools as command-line tools
-------------------------------------

To use python-oletools from the command line as analysis tools, you may simply 
[download the latest release archive](https://github.com/decalage2/oletools/releases)
and extract the files into the directory of your choice.

You may also download the [latest development version](https://github.com/decalage2/oletools/archive/master.zip) with the most recent features.

Another possibility is to use a git client to clone the repository (https://github.com/decalage2/oletools.git) into a folder.
You can then update it easily in the future.

### Windows

You may add the oletools directory to your PATH environment variable to access the tools from anywhere.

### Linux, Mac OSX, Unix

It is very convenient to create symbolic links to each tool in one of the bin directories in order to run them as shell 
commands from anywhere. For example, here is how to create an executable link "olevba" in `/usr/local/bin` pointing to
olevba.py, assuming oletools was unzipped into /opt/oletools:

```text
chmod +x /opt/oletools/oletools/olevba.py
ln -s /opt/oletools/oletools/olevba.py /usr/local/bin/olevba
```
Then the olevba command can be used from any directory:

```text
user@remnux:~/MalwareZoo/VBA$ olevba dridex427.xls |less
```

For python applications
-----------------------

If you plan to use python-oletools with other Python applications or your own scripts, the simplest solution is to use 
**"pip install oletools"** or **"easy_install oletools"** to download and install the package in one go. Pip is included
with Python since version 2.7.9.

**Important: to update oletools** if it is already installed, you must run **"pip install -U oletools"**, otherwise pip
will not update it.

Alternatively if you prefer the old school way, you may download the 
[latest archive](https://github.com/decalage2/oletools/releases), extract it into
a temporary directory and run **"python setup.py install"**.

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
