' Writes expected.tsv: each expression of expressions.txt evaluated by VBScript, whose operators
' and conversions are VBA's. Run with: cscript //nologo make_expected.vbs
' No PowerPoint, no UI: errors are caught and written as ERROR <number>.
Option Explicit
Dim fso, here, inFile, outFile, line, x, p, q, r, h, z, s, t, u
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
p = 7 : q = 2 : r = -7.5 : h = 2.5 : z = 0 : s = "10" : t = "ab" : u = True

Set inFile = fso.OpenTextFile(here & "\expressions.txt", 1)
Set outFile = fso.CreateTextFile(here & "\expected.tsv", True)
Do Until inFile.AtEndOfStream
    line = Trim(inFile.ReadLine)
    If Len(line) > 0 And Left(line, 1) <> "#" Then
        On Error Resume Next
        Err.Clear
        Execute "x = " & line
        If Err.Number <> 0 Then
            outFile.WriteLine line & vbTab & "ERROR" & vbTab & Err.Number
        Else
            outFile.WriteLine line & vbTab & TypeName(x) & vbTab & CStr(x)
        End If
        On Error GoTo 0
    End If
Loop
inFile.Close
outFile.Close
WScript.Echo "expected.tsv written"
