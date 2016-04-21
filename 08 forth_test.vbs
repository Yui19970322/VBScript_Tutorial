Option Explicit

Dim i

For i = 1 To 5
	MsgBox i & " 回目", vbInformation, WScript.ScriptName
	'もしも、最後の数字だったら、「最後の数字です」と表示する。
	
	If 	i= 5 Then
		MsgBox "最後の数字です", vbQuestion, WScript.ScriptName
	End If
Next

