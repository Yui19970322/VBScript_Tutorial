Option Explicit

Dim i
Dim Result

'1から10までの数字を表示する。
For i = 1 To 10
    WScript.StdOut.Write i & " "
Next

WScript.StdOut.Write vbCrLf


'Result = For i = 1 To 10
'1から10までの数字の中の奇数だけを表示する。
 For i = 1 To 10

    Result = i Mod 2

    If Result = 1 Then
        WScript.StdOut.Write i & " "
    End If
Next



WScript.StdOut.Write vbCrLf
'1から10までの数字の中の偶数だけを表示する。
For i = 1 To 10

    Result = i Mod 2

    If Result = 0 Then
        WScript.StdOut.Write i & " "
    End If
Next

