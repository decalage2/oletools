' DocVarDump.vba
'
' DocVarDump is a VBA macro that can be used to dump the content of all document
' variables stored in a MS Word document.
'
' USAGE:
'  1. Open the document to be analyzed in MS Word
'  2. Do NOT click on "Enable Content", to avoid running malicious macros
'  3. Save the document with a new name, using the DOCX format (not doc, not docm)
'     This will remove all VBA macro code.
'  4. Close the file, and reopen the DOCX file you just saved
'  5. Press Alt+F11 to open the VBA Editor
'  6. Double-click on "This Document" under Project
'  7. Copy and Paste all the code from DocVarDump.vba
'  8. Move the cursor on the line "Sub DocVarDump()"
'  9. Press F5: This should run the code, and create a file "docvardump.txt"
'     containing a hex dump of all document variables.
'
' ALTERNATIVE: Open the document in LibreOffice/OpenOffice,
' then go to File / Properties / Custom Properties
'
' Author: Philippe Lagadec - http://www.decalage.info
' License: BSD, see source code or documentation
'
' DocVarDump is part of the python-oletools package:
' http://www.decalage.info/python/oletools

' CHANGELOG:
' 2016-09-21 v0.01 PL: - First working version
' 2017-04-10 v0.02 PL: - Added usage instructions

Sub DocVarDump()
    intFileNum = FreeFile
    FName = Environ("TEMP") & "\docvardump.txt"
    Open FName For Output As intFileNum
        For Each myvar In ActiveDocument.Variables
            Write #intFileNum, "Name = " & myvar.Name
            'TODO: check VarType, and only use hexdump for strings with non-printable chars
            Write #intFileNum, "Value = " & HexDump(myvar.value)
            Write #intFileNum,
        Next myvar
    Close intFileNum
    Documents.Open (FName)
End Sub

Function Hex2(value As Integer)
    h = Hex(value)
    If Len(h) < 2 Then
        h = "0" & h
    End If
    Hex2 = h
End Function

Function HexN(value As Integer, nchars As Integer)
    h = Hex(value)
    Do While Len(h) < nchars
        h = "0" & h
    Loop
    HexN = h
End Function

Function ReplaceClean1(sText As String)
    Dim J As Integer
    Dim vAddText

    vAddText = Array(Chr(129), Chr(141), Chr(143), Chr(144), Chr(157))
    For J = 0 To 31
        sText = Replace(sText, Chr(J), "\x" & Hex2(J))
    Next
    For J = 0 To UBound(vAddText)
        c = vAddText(J)
        a = Asc(c)
        sText = Replace(sText, c, "\x" & Hex2(a))
    Next
    ReplaceClean1 = sText
End Function

Function ReplaceClean3(sText As String)
    Dim J As Integer
    For J = 0 To 31
        sText = Replace(sText, Chr(J), ".")
    Next
    For J = 127 To 255
        sText = Replace(sText, Chr(J), ".")
    Next
    ReplaceClean3 = sText
End Function

Function HexBytes(sText As String)
    Dim i As Integer
    HexBytes = ""
    For i = 1 To Len(sText)
        HexBytes = HexBytes & Hex2(Asc(Mid(sText, i))) & " "
    Next
End Function


Function HexDump(sText As String)
    Dim chunk As String
    Dim i As Long
    ' "\" is integer division, "/" is normal division (float)
    nbytes = 8
    nchunks = Len(sText) \ nbytes
    lastchunk = Len(sText) Mod nbytes
    HexDump = ""
    For i = 0 To nchunks - 1
        Offset = HexN(i * nbytes, 8)
        chunk = Mid(sText, i * nbytes + 1, nbytes)
        HexDump = HexDump & Offset & "  " & HexBytes(chunk) & " " & ReplaceClean3(chunk) & vbCrLf
    Next i
    'TODO: LAST CHUNK!
    If lastchunk > 0 Then
        Offset = HexN(nchunks * nbytes, 8)
        chunk = Mid(sText, nchunks * nbytes + 1, lastchunk)
        HexDump = HexDump & Offset & "  " & HexBytes(chunk) & " " & ReplaceClean3(chunk) & vbCrLf
    End If
End Function
