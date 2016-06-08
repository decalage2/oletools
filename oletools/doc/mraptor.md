mraptor (MacroRaptor)
=====================

mraptor is a script to detect malicious VBA Macros.

It can be used either as a command-line tool, or as a python module from your own applications.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

## Usage

```text
Usage: mraptor.py [options] <filename> [filename2 ...]

Options:
  -h, --help            show this help message and exit
  -r                    find files recursively in subdirectories.
  -z ZIP_PASSWORD, --zip=ZIP_PASSWORD
                        if the file is a zip archive, open all files from it,
                        using the provided password (requires Python 2.6+)
  -f ZIP_FNAME, --zipfname=ZIP_FNAME
                        if the file is a zip archive, file(s) to be opened
                        within the zip. Wildcards * and ? are supported.
                        (default:*)
  -l LOGLEVEL, --loglevel=LOGLEVEL
                        logging level debug/info/warning/error/critical
                        (default=warning)
  -m, --matches         Show matched strings.

An exit code is returned based on the analysis result:
 - 0: No Macro
 - 1: Not MS Office
 - 2: Macro OK
 - 10: ERROR
 - 20: SUSPICIOUS
```

### Examples

Scan a single file:

```text
mraptor.py file.doc
```

Scan a single file, stored in a Zip archive with password "infected":

```text
mraptor.py malicious_file.xls.zip -z infected
```

Scan a collection of files stored in a folder:

```text
mraptor.py "MalwareZoo/VBA/*"
```

**Important**: on Linux/MacOSX, always add double quotes around a file name when you use
wildcards such as `*` and `?`. Otherwise, the shell may replace the argument with the actual
list of files matching the wildcards before starting the script.

![](mraptor1.png)

--------------------------------------------------------------------------
    
## How to use mraptor in Python applications

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
	- [[oledir]]
	- [[olemap]]
	- [[olevba]]
	- [[mraptor]]
	- [[pyxswf]]
	- [[oleobj]]
	- [[rtfobj]]
