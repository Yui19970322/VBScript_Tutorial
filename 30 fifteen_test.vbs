Option Explicit

Dim i
Dim Result

'1����10�܂ł̐�����\������B
For i = 1 To 10
    WScript.StdOut.Write i & " "
Next

WScript.StdOut.Write vbCrLf


'Result = For i = 1 To 10
'1����10�܂ł̐����̒��̊������\������B
 For i = 1 To 10

    Result = i Mod 2

    If Result = 1 Then
        WScript.StdOut.Write i & " "
    End If
Next



WScript.StdOut.Write vbCrLf
'1����10�܂ł̐����̒��̋���������\������B
For i = 1 To 10

    Result = i Mod 2

    If Result = 0 Then
        WScript.StdOut.Write i & " "
    End If
Next

