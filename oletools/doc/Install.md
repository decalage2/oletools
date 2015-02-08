How to Download and Install python-oletools
===========================================

Pre-requisites
--------------

For now, python-oletools require Python 2.x, if possible 2.6 or 2.7. They are not compatible with Python 3.x yet.


To use oletools as command-line tools
-------------------------------------

To use python-oletools from the command line as analysis tools, you may simply 
[download the zip archive](https://bitbucket.org/decalage/oletools/downloads) 
and extract the files in the directory of your choice. Pick the latest release version, or click on "Download Repository"
to get the latest development version with the most recent features.

Another possibility is to use a Mercurial client (hg) to clone the repository in a folder. You can then update it easily
in the future.

You may add the oletools directory to your PATH environment variable to access the tools from anywhere.


For python applications
-----------------------

If you plan to use python-oletools with other Python applications or your own scripts, the simplest solution is to use 
**"pip install oletools"** or **"easy_install oletools"** to download and install the package in one go. Pip is included
with Python since version 2.7.9.

**Important: to update oletools** if it is already installed, you must run **"pip install -U oletools"**, otherwise pip
will not update it.

Alternatively, you may download/extract the [zip archive](https://bitbucket.org/decalage/oletools/downloads) in a temporary 
directory and run **"python setup.py install"**.

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