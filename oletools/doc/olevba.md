olevba
======

olevba is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to **detect VBA Macros**, extract their **source code** in clear text, 
and detect security-related patterns such as **auto-executable macros**, **suspicious
VBA keywords** used by malware, anti-sandboxing and anti-virtualization techniques, 
and potential **IOCs** (IP addresses, URLs, executable filenames, etc). 
It also detects and decodes several common **obfuscation methods including Hex encoding,
StrReverse, Base64, Dridex, VBA expressions**, and extracts IOCs from decoded strings.
XLM/Excel 4 Macros are also supported in Excel and SLK files.

It can be used either as a command-line tool, or as a python module from your own applications.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

olevba is based on source code from [officeparser](https://github.com/unixfreak0037/officeparser) 
by John William Davison, with significant modifications.

## Supported formats

- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 97-2003 (.ppt), PowerPoint 2007+ (.pptm, .ppsm)
- Word/PowerPoint 2007+ XML (aka Flat OPC)
- Word 2003 XML (.xml)
- Word/Excel Single File Web Page / MHTML (.mht)
- Publisher (.pub)
- SYLK/SLK files (.slk)
- Text file containing VBA or VBScript source code
- Password-protected Zip archive containing any of the above

S## Main Features

- Detect VBA macros in MS Office 97-2003 and 2007+ files, XML, MHT
- Extract VBA macro source code
- Detect auto-executable macros
- Detect suspicious VBA keywords often used by malware
- Detect anti-sandboxing and anti-virtualization techniques
- Detect and decodes strings obfuscated with Hex/Base64/StrReverse/Dridex
- Deobfuscates VBA expressions with any combination of Chr, Asc, Val, StrReverse, Environ, +, &, using a VBA parser built with
[pyparsing](http://pyparsing.wikispaces.com), including custom Hex and Base64 encodings
- Extract IOCs/patterns of interest such as IP addresses, URLs, e-mail addresses and executable file names
- Scan multiple files and sample collections (wildcards, recursive)
- Triage mode for a summary view of multiple files
- Scan malware samples in password-protected Zip archives
- Python API to use olevba from your applications

MS Office files encrypted with a password are also supported, because VBA macro code is never
encrypted, only the content of the document.

## About VBA Macros

See [this article](http://www.decalage.info/en/vba_tools) for more information and technical details about VBA Macros
and how they are stored in MS Office documents.

## How it works

1. olevba checks the file type: If it is an OLE file (i.e MS Office 97-2003), it is parsed right away.
1. If it is a zip file (i.e. MS Office 2007+), XML or MHTML, olevba looks for all OLE files stored in it (e.g. vbaProject.bin, editdata.mso), and opens them.
1. olevba identifies all the VBA projects stored in the OLE structure.
1. Each VBA project is parsed to find the corresponding OLE streams containing macro code.
1. In each of these OLE streams, the VBA macro source code is extracted and decompressed (RLE compression).
1. olevba looks for specific strings obfuscated with various algorithms (Hex, Base64, StrReverse, Dridex, VBA expressions).
1. olevba scans the macro source code and the deobfuscated strings to find suspicious keywords, auto-executable macros
and potential IOCs (URLs, IP addresses, e-mail addresses, executable filenames, etc).


## Usage

```text
Usage: olevba [options] <filename> [filename2 ...]

Options:
  -h, --help            show this help message and exit
  -r                    find files recursively in subdirectories.
  -z ZIP_PASSWORD, --zip=ZIP_PASSWORD
                        if the file is a zip archive, open all files from it,
                        using the provided password.
  -p PASSWORD, --password=PASSWORD
                        if encrypted office files are encountered, try
                        decryption with this password. May be repeated.
  -f ZIP_FNAME, --zipfname=ZIP_FNAME
                        if the file is a zip archive, file(s) to be opened
                        within the zip. Wildcards * and ? are supported.
                        (default:*)
  -a, --analysis        display only analysis results, not the macro source
                        code
  -c, --code            display only VBA source code, do not analyze it
  --decode              display all the obfuscated strings with their decoded
                        content (Hex, Base64, StrReverse, Dridex, VBA).
  --attr                display the attribute lines at the beginning of VBA
                        source code
  --reveal              display the macro source code after replacing all the
                        obfuscated strings by their decoded content.
  -l LOGLEVEL, --loglevel=LOGLEVEL
                        logging level debug/info/warning/error/critical
                        (default=warning)
  --deobf               Attempt to deobfuscate VBA expressions (slow)
  --relaxed             Do not raise errors if opening of substream fails

  Output mode (mutually exclusive):
    -t, --triage        triage mode, display results as a summary table
                        (default for multiple files)
    -d, --detailed      detailed mode, display full results (default for
                        single file)
    -j, --json          json mode, detailed in json format (never default)
```

**New in v0.54:** the -p option can now be used to decrypt encrypted documents using the provided password(s).

### Examples

Scan a single file:

```text
olevba file.doc
```
    
Scan a single file, stored in a Zip archive with password "infected":

```text
olevba malicious_file.xls.zip -z infected
```
    
Scan a single file, showing all obfuscated strings decoded:

```text
olevba file.doc --decode
```
    
Scan a single file, showing the macro source code with VBA strings deobfuscated:

```text
olevba file.doc --reveal
```

Scan VBA source code extracted into a text file:

```text
olevba source_code.vba
```

Scan a collection of files stored in a folder:

```text
olevba "MalwareZoo/VBA/*"
```
NOTE: On Linux, MacOSX and other Unix variants, it is required to add double quotes around wildcards. Otherwise, they will be expanded by the shell instead of olevba.

Scan all .doc and .xls files, recursively in all subfolders:

```text
olevba "MalwareZoo/VBA/*.doc" "MalwareZoo/VBA/*.xls" -r
```

Scan all .doc files within all .zip files with password, recursively:

```text
olevba "MalwareZoo/VBA/*.zip" -r -z infected -f "*.doc"
```


### Detailed analysis mode (default for single file)

When a single file is scanned, or when using the option -d, all details of the analysis are displayed. 

For example, checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

```text
>olevba c:\MalwareZoo\VBA\DIAN_caso-5415.doc.zip -z infected
===============================================================================
FILE: DIAN_caso-5415.doc.malware in c:\MalwareZoo\VBA\DIAN_caso-5415.doc.zip
Type: OLE
-------------------------------------------------------------------------------
VBA MACRO ThisDocument.cls
in file: DIAN_caso-5415.doc.malware - OLE stream: Macros/VBA/ThisDocument
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
Private Declare Function URLDownloadToFileA Lib "urlmon" (ByVal FVQGKS As Long,_
ByVal WSGSGY As String, ByVal IFRRFV As String, ByVal NCVOLV As Long, _
ByVal HQTLDG As Long) As Long
Sub AutoOpen()
    Auto_Open
End Sub
Sub Auto_Open()
SNVJYQ
End Sub
Public Sub SNVJYQ()
    [Malicious Code...]
End Sub
Function OGEXYR(XSTAHU As String, PHHWIV As String) As Boolean
    [Malicious Code...]
    Application.DisplayAlerts = False
    Application.Quit
End Function
Sub Workbook_Open()
    Auto_Open
End Sub

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ANALYSIS:
+------------+----------------------+-----------------------------------------+
| Type       | Keyword              | Description                             |
+------------+----------------------+-----------------------------------------+
| AutoExec   | AutoOpen             | Runs when the Word document is opened   |
| AutoExec   | Auto_Open            | Runs when the Excel Workbook is opened  |
| AutoExec   | Workbook_Open        | Runs when the Excel Workbook is opened  |
| Suspicious | Lib                  | May run code from a DLL                 |
| Suspicious | Shell                | May run an executable file or a system  |
|            |                      | command                                 |
| Suspicious | Environ              | May read system environment variables   |
| Suspicious | URLDownloadToFileA   | May download files from the Internet    |
| IOC        | http://germanya.com. | URL                                     |
|            | ec/logs/test.exe"    |                                         |
| IOC        | http://germanya.com. | URL                                     |
|            | ec/logs/counter.php" |                                         |
| IOC        | germanya.com         | Executable file name                    |
| IOC        | test.exe             | Executable file name                    |
| IOC        | sfjozjero.exe        | Executable file name                    |
+------------+----------------------+-----------------------------------------+
```

### Triage mode (default for multiple files)

When several files are scanned, or when using the option -t, a summary of the analysis for each file is displayed.
This is more convenient for quick triage of a collection of suspicious files.

The following flags show the results of the analysis:

- **OLE**: the file type is OLE, for example MS Office 97-2003
- **OpX**: the file type is OpenXML, for example MS Office 2007+
- **XML**: the file type is Word 2003 XML
- **MHT**: the file type is Word MHTML, aka Single File Web Page (.mht)
- **?**: the file type is not supported
- **M**: contains VBA Macros
- **A**: auto-executable macros
- **S**: suspicious VBA keywords
- **I**: potential IOCs
- **H**: hex-encoded strings (potential obfuscation)
- **B**: Base64-encoded strings (potential obfuscation)
- **D**: Dridex-encoded strings (potential obfuscation)
- **V**: VBA string expressions (potential obfuscation)

Here is an example:

```text
c:\>olevba \MalwareZoo\VBA\samples\*
Flags       Filename
----------- -----------------------------------------------------------------
OLE:MASI--- \MalwareZoo\VBA\samples\DIAN_caso-5415.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_1.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_2.doc.malware
OLE:MASI--- \MalwareZoo\VBA\samples\DRIDEX_3.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_4.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_5.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_6.doc.malware
OLE:MAS---- \MalwareZoo\VBA\samples\DRIDEX_7.doc.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_8.doc.malware
OLE:MASIHBD \MalwareZoo\VBA\samples\DRIDEX_9.xls.malware
OLE:MASIH-- \MalwareZoo\VBA\samples\DRIDEX_A.doc.malware
OLE:------- \MalwareZoo\VBA\samples\Normal_Document.doc
OLE:M------ \MalwareZoo\VBA\samples\Normal_Document_Macro.doc
OpX:MASI--- \MalwareZoo\VBA\samples\RottenKitten.xlsb.malware
OLE:MASI-B- \MalwareZoo\VBA\samples\ROVNIX.doc.malware
OLE:MA----- \MalwareZoo\VBA\samples\Word within Word macro auto.doc
```

## Python 3 support - olevba3

Since v0.54, olevba is fully compatible with both Python 2 and 3.
There is no need to use olevba3 anymore, however it is still present for backward compatibility.


--------------------------------------------------------------------------
    
## How to use olevba in Python applications	

olevba may be used to open a MS Office file, detect if it contains VBA macros, extract and analyze the VBA source code 
from your own python applications.

IMPORTANT: olevba is currently under active development, therefore this API is likely to change.

### Import olevba

First, import the **oletools.olevba** package, using at least the VBA_Parser and VBA_Scanner classes:

```python
from oletools.olevba import VBA_Parser, TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML, TYPE_MHTML
```

### Parse a MS Office file - VBA_Parser

To parse a file on disk, create an instance of the **VBA_Parser** class, providing the name of the file to open as parameter.
For example:

```python
vbaparser = VBA_Parser('my_file_with_macros.doc')
```

The file may also be provided as a bytes string containing its data. In that case, the actual 
filename must be provided for reference, and the file content with the data parameter. For example:

```python
myfile = 'my_file_with_macros.doc'
filedata = open(myfile, 'rb').read()
vbaparser = VBA_Parser(myfile, data=filedata)
```
VBA_Parser will raise an exception if the file is not a supported format, such as OLE (MS Office 97-2003), OpenXML 
(MS Office 2007+), MHTML or Word 2003 XML.
 
After parsing the file, the attribute **VBA_Parser.type** is a string indicating the file type.
It can be either TYPE_OLE, TYPE_OpenXML, TYPE_Word2003_XML or TYPE_MHTML. (constants defined in the olevba module)

### Detect VBA macros

The method **detect_vba_macros** of a VBA_Parser object returns True if VBA macros have been found in the file, 
False otherwise.

```python
if vbaparser.detect_vba_macros():
    print 'VBA Macros found'
else:
    print 'No VBA Macros found'
```
Note: The detection algorithm looks for streams and storage with specific names in the OLE structure, which works fine
for all the supported formats listed above. However, for some formats such as PowerPoint 97-2003, this method will 
always return False because VBA Macros are stored in a different way which is not yet supported by olevba.

Moreover, if the file contains an embedded document (e.g. an Excel workbook inserted into a Word document), this method
may return True if the embedded document contains VBA Macros, even if the main document does not.
 
### Extract VBA Macro Source Code
 
The method **extract_macros** extracts and decompresses source code for each VBA macro found in the file (possibly 
including embedded files). It is a generator yielding a tuple (filename, stream_path, vba_filename, vba_code) 
for each VBA macro found.

- filename: If the file is OLE (MS Office 97-2003), filename is the path of the file.
    If the file is OpenXML (MS Office 2007+), filename is the path of the OLE subfile containing VBA macros within the zip archive, 
    e.g. word/vbaProject.bin.
- stream_path: path of the OLE stream containing the VBA macro source code
- vba_filename: corresponding VBA filename
- vba_code: string containing the VBA source code in clear text

Example:

```python
for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
    print '-'*79
    print 'Filename    :', filename
    print 'OLE stream  :', stream_path
    print 'VBA filename:', vba_filename
    print '- '*39
    print vba_code
```
Alternatively, the VBA_Parser method **extract_all_macros** returns the same results as a list of tuples.

### Analyze VBA Source Code

Since version 0.40, the VBA_Parser class provides simpler methods than VBA_Scanner to analyze all macros contained
in a file:

The method **analyze_macros** from the class **VBA_Parser** can be used to scan the source code of all
VBA modules to find obfuscated strings, suspicious keywords, IOCs, auto-executable macros, etc.

analyze_macros() takes an optional argument show_decoded_strings: if set to True, the results will contain all the encoded
strings found in the code (Hex, Base64, Dridex) with their decoded value.
By default, it will only include the strings which contain printable characters.

**VBA_Parser.analyze_macros()** returns a list of tuples (type, keyword, description), one for each item in the results.

- type may be either 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String', 'Dridex String' or 
  'VBA obfuscated Strings'.
- keyword is the string found for auto-executable macros, suspicious keywords or IOCs. For obfuscated strings, it is
  the decoded value of the string.
- description provides a description of the keyword. For obfuscated strings, it is the encoded value of the string.

Example:

```python
results = vbaparser.analyze_macros()
for kw_type, keyword, description in results:
    print 'type=%s - keyword=%s - description=%s' % (kw_type, keyword, description)
```
After calling analyze_macros, the following VBA_Parser attributes also provide the number
of items found for each category:

```python
print 'AutoExec keywords: %d' % vbaparser.nb_autoexec
print 'Suspicious keywords: %d' % vbaparser.nb_suspicious
print 'IOCs: %d' % vbaparser.nb_iocs
print 'Hex obfuscated strings: %d' % vbaparser.nb_hexstrings
print 'Base64 obfuscated strings: %d' % vbaparser.nb_base64strings
print 'Dridex obfuscated strings: %d' % vbaparser.nb_dridexstrings
print 'VBA obfuscated strings: %d' % vbaparser.nb_vbastrings
```

### Deobfuscate VBA Macro Source Code

The method **reveal** attempts to deobfuscate the macro source code by replacing all
the obfuscated strings by their decoded content. Returns a single string.

Example:

```python
print vbaparser.reveal()
```

### Close the VBA_Parser

After usage, it is better to call the **close** method of the VBA_Parser object, to make sure the file is closed, 
especially if your application is parsing many files.

```python
vbaparser.close()
```

--------------------------------------------------------------------------

## Deprecated API

The following methods and functions are still functional, but their usage is not recommended
since they have been replaced by better solutions.

### VBA_Scanner (deprecated)

The class **VBA_Scanner** can be used to scan the source code of a VBA module to find obfuscated strings,
suspicious keywords, IOCs, auto-executable macros, etc.

First, create a VBA_Scanner object with a string containing the VBA source code (for example returned by the 
extract_macros method). Then call the methods **scan** or **scan_summary** to get the results of the analysis.

scan() takes an optional argument include_decoded_strings: if set to True, the results will contain all the encoded
strings found in the code (Hex, Base64, Dridex) with their decoded value.

**scan** returns a list of tuples (type, keyword, description), one for each item in the results. 

- type may be either 'AutoExec', 'Suspicious', 'IOC', 'Hex String', 'Base64 String' or 'Dridex String'.
- keyword is the string found for auto-executable macros, suspicious keywords or IOCs. For obfuscated strings, it is
  the decoded value of the string.
- description provides a description of the keyword. For obfuscated strings, it is the encoded value of the string.

Example:

```python
vba_scanner = VBA_Scanner(vba_code)
results = vba_scanner.scan(include_decoded_strings=True)
for kw_type, keyword, description in results:
    print 'type=%s - keyword=%s - description=%s' % (kw_type, keyword, description)
```
The function **scan_vba** is a shortcut for VBA_Scanner(vba_code).scan():

```python
results = scan_vba(vba_code, include_decoded_strings=True)
for kw_type, keyword, description in results:
    print 'type=%s - keyword=%s - description=%s' % (kw_type, keyword, description)
```
**scan_summary** returns a tuple with the number of items found for each category: 
(autoexec, suspicious, IOCs, hex, base64, dridex).


### Detect auto-executable macros (deprecated)

**Deprecated**: It is preferable to use either scan_vba or VBA_Scanner to get all results at once. 

The function **detect_autoexec** checks if VBA macro code contains specific macro names
that will be triggered when the document/workbook is opened, closed, changed, etc.

It returns a list of tuples containing two strings, the detected keyword, and the
description of the trigger. (See the malware example above)

Sample usage:

```python
from oletools.olevba import detect_autoexec
autoexec_keywords = detect_autoexec(vba_code)
if autoexec_keywords:
    print 'Auto-executable macro keywords found:'
    for keyword, description in autoexec_keywords:
        print '%s: %s' % (keyword, description)
else:
    print 'Auto-executable macro keywords: None found'
```

### Detect suspicious VBA keywords (deprecated)

**Deprecated**: It is preferable to use either scan_vba or VBA_Scanner to get all results at once. 

The function **detect_suspicious** checks if VBA macro code contains specific
keywords often used by malware to act on the system (create files, run
commands or applications, write to the registry, etc).

It returns a list of tuples containing two strings, the detected keyword, and the
description of the corresponding malicious behaviour. (See the malware example above)

Sample usage:

```python
from oletools.olevba import detect_suspicious
suspicious_keywords = detect_suspicious(vba_code)
if suspicious_keywords:
    print 'Suspicious VBA keywords found:'
    for keyword, description in suspicious_keywords:
        print '%s: %s' % (keyword, description)
else:
    print 'Suspicious VBA keywords: None found'
```

### Extract potential IOCs (deprecated)

**Deprecated**: It is preferable to use either scan_vba or VBA_Scanner to get all results at once. 

The function **detect_patterns** checks if VBA macro code contains specific
patterns of interest, that may be useful for malware analysis and detection
(potential Indicators of Compromise): IP addresses, e-mail addresses,
URLs, executable file names.

It returns a list of tuples containing two strings, the pattern type, and the
extracted value. (See the malware example above)

Sample usage:

```python
from oletools.olevba import detect_patterns
patterns = detect_patterns(vba_code)
if patterns:
    print 'Patterns found:'
    for pattern_type, value in patterns:
        print '%s: %s' % (pattern_type, value)
else:
    print 'Patterns: None found'
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
