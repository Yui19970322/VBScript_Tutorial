Option Explicit

Dim objWShell
Set objWShell = Createobject("WScript.Shell")

'���[�U�[���琔�l�̓��͂��󂯕t���܂��B
'���[�U�[�̓��͂������l������������𔻒肵�܂��B
'���[�U�[�̓��͂������l�������̎��́APopup��3�b�ԕ\�����܂��B
'���[�U�[�̓��͂������l����̎��́APopup�����[�U�[�̓��͂������l�̕b�������\�����܂��B





objWShell.Popup "���͂悤�������܂�" _ 
              , 5 _
              , "Popup�T���v��" _
              ,vbOkOnly + vbInformation


Set objWShell = Nothing
