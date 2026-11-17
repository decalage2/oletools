How to Download and Install oletools
====================================

Pre-requisites
--------------

The recommended Python version to run oletools is the latest **Python 3.x** (3.12 for now). 
Python 2.7 is still supported for the moment, even if it reached end of life in 2020 
(for projects still using Python 2/PyPy 2 such as ViperMonkey).
It is highly recommended to switch to Python 3 if possible.

Recommended way to Download+Install/Update oletools: pip or pipx
----------------------------------------------------------------

Pip is included with Python since version 2.7.9 and 3.4. If it is not installed on your
system, either upgrade Python or see https://pip.pypa.io/en/stable/installing/

### Linux, Mac OSX, Unix

To download and install/update the latest release version of oletools with all its dependencies,
run the following command in a shell:

```text
sudo -H pip install -U oletools[full]
```
The keyword `[full]` means that all optional dependencies will be installed, such as XLMMacroDeobfuscator.
If you prefer a lighter version without optional dependencies, use the following command instead:

```text
sudo -H pip install -U oletools
```

Replace `pip` by `pip3` or `pip2` to install on a specific Python version.

On some Linux distributions, it might not be allowed to install system-wide python packages
with pip. In that case, pipx may be a better alternative to install oletools in a user virtual
environment, and to install the command-line scripts oleid, olevba, etc:

```text
pipx install oletools
```


**Important**: Since version 0.50, pip will automatically create convenient command-line scripts
in /usr/local/bin to run all the oletools from any directory.

### Windows

To download and install/update the latest release version of oletools with all its dependencies,
run the following command in a cmd window:

```text
pip install -U oletools[full]
```
The keyword `[full]` means that all optional dependencies will be installed, such as XLMMacroDeobfuscator.
If you prefer a lighter version without optional dependencies, use the following command instead:

```text
pip install -U oletools
```

Replace `pip` by `pip3` or `pip2` to install on a specific Python version.

**Note**: with Python 3, you may need to open a cmd window with Administrator privileges in order to run pip 
and install for all users. If that is not possible, you may also install only for the current user 
by adding the `--user` option:

```text
pip3 install -U --user oletools
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
Note that it will install oletools without optional dependencies such as XLMMacroDeobfuscator,
so you may need to install them separately.

Replace `pip` by `pip3` or `pip2` to install on a specific Python version.

### Windows

```text
pip install -U https://github.com/decalage2/oletools/archive/master.zip
```
Note that it will install oletools without optional dependencies such as XLMMacroDeobfuscator,
so you may need to install them separately.

Replace `pip` by `pip3` or `pip2` to install on a specific Python version.

**Note**: with Python 3, you may need to open a cmd window with Administrator privileges in order to run pip 
and install for all users. If that is not possible, you may also install only for the current user 
by adding the `--user` option:

```text
pip3 install -U --user https://github.com/decalage2/oletools/archive/master.zip
```

How to install offline - Computer without Internet access
---------------------------------------------------------

First, download the oletools archive on a computer with Internet access:
* Latest stable version: from https://pypi.org/project/oletools/ or https://github.com/decalage2/oletools/releases
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
