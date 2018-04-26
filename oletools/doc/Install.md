How to Download and Install python-oletools
===========================================

Pre-requisites
--------------

The recommended Python version to run oletools is **Python 2.7**.
Python 2.6 is also supported, but as it is not tested as often as 2.7, some features
might not work as expected.

Since oletools v0.50, thanks to contributions by [@Sebdraven](https://twitter.com/Sebdraven),
most tools can also run with **Python 3.x**. As this is quite new, please
[report any issue]((https://github.com/decalage2/oletools/issues)) you may encounter.



Recommended way to Download+Install/Update oletools: pip
--------------------------------------------------------

Pip is included with Python since version 2.7.9 and 3.4. If it is not installed on your
system, either upgrade Python or see https://pip.pypa.io/en/stable/installing/

### Linux, Mac OSX, Unix

To download and install/update the latest release version of oletools,
run the following command in a shell:

```text
sudo -H pip install -U oletools
```

**Important**: Since version 0.50, pip will automatically create convenient command-line scripts
in /usr/local/bin to run all the oletools from any directory.

### Windows

To download and install/update the latest release version of oletools,
run the following command in a cmd window:

```text
pip install -U oletools
```

**Important**: Since version 0.50, pip will automatically create convenient command-line scripts
to run all the oletools from any directory: olevba, mraptor, oleid, rtfobj, etc.


How to install the latest development version
---------------------------------------------

If you want to benefit from the latest improvements in the development version,
you may also use pip:

### Linux, Mac OSX, Unix

```text
sudo -H pip install -U https://github.com/decalage2/oletools/archive/master.zip
```

### Windows

```text
pip install -U https://github.com/decalage2/oletools/archive/master.zip
```

How to install offline - Computer without Internet access
---------------------------------------------------------

First, download the oletools archive on a computer with Internet access:
* Latest stable version: from https://github.com/decalage2/oletools/releases
* Development version: https://github.com/decalage2/oletools/archive/master.zip

Copy the archive file to the target computer.

On Linux, Mac OSX, Unix, run the following command using the filename of the
archive that you downloaded:

```text
sudo -H pip install -U oletools.zip
```

On Windows:

```text
pip install -U oletools.zip
```


Old school install using setup.py
---------------------------------

If you cannot use pip, it is still possible to run the setup.py script
directly. However, this method will not create the command-line scripts
automatically.

First, download the oletools archive:
* Latest stable version: from https://github.com/decalage2/oletools/releases
* Development version: https://github.com/decalage2/oletools/archive/master.zip

Then extract the archive, open a shell and go to the oletools directory.

### Linux, Mac OSX, Unix

```text
sudo -H python setup.py install
```

### Windows:

```text
python setup.py install
```


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
