Option Explicit

'�R�}���h�v�����v�g�Ŏ��s�����Ƃ�����
'Hello! �R�}���h�v�����v�g�ƕ\��������
'WScript.echo "Hello! �R�}���h�v�����v�g"


Dim strWSH

strWSH = UCase(Right(WScript.FullName,11))

Select Case strWSH
       
       Case "WSCRIPT.EXE"
            MsgBox "GUI�Ŏ��s����܂���!" _
                 , vbOkOnly + vbInformation _
                 , "���s���[�h"

       Case "CSCRIPT.EXE"
            WScript.Echo "Hello! �R�}���h�v�����v�g"

End Select