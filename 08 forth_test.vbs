Option Explicit

Dim i

For i = 1 To 5
	MsgBox i & " ���", vbInformation, WScript.ScriptName
	'�������A�Ō�̐�����������A�u�Ō�̐����ł��v�ƕ\������B
	
	If 	i= 5 Then
		MsgBox "�Ō�̐����ł�", vbQuestion, WScript.ScriptName
	End If
Next

