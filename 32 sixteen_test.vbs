Option Explicit

Dim Number
Dim objWShell
Dim Result
Set objWShell = Createobject("WScript.Shell")


'���[�U�[���琔�l�̓��͂��󂯕t���܂��B
WScript.Echo "�������l����͂��Ă��������B"
Number = WScript.StdIn.ReadLine
WScript.Echo Number & "�����͂���܂����B"

'���[�U�[�̓��͂������l������������𔻒肵�܂��B
Result = CInt(Number) Mod 2
If Result = 0 Then
    WScript.Echo Number & "�����ł�"
ElseIf Result = 1  Then
    WScript.Echo Number & "��ł�"
End IF


'���[�U�[�̓��͂������l�������̎��́APopup��3�b�ԕ\�����܂��B
If Result = 0 Then
    objWShell.Popup "����ɂ���" _ 
              , 3 _
              , "Popup�T���v��" _
              ,vbOkOnly + vbInformation
End If             

'���[�U�[�̓��͂������l����̎��́APopup�����[�U�[�̓��͂������l�̕b�������\�����܂��B
If Result = 1 Then
    objWShell.Popup "����΂��" _ 
              ,(Number) _ 
              , "Popup�T���v��" _
              ,vbOkOnly + vbInformation
End If             

Set objWShell = Nothing
