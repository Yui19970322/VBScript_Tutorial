Option Explicit

Dim strOrg
Dim strKey
Dim strNew

strOrg = "�����ŗ�����E���ł�������"
strKey = "����"

'���[�U�[����̓��͂��󂯕t���܂��B
'"�����ŗ�����E���ł�������" "����" ���A���[�U�[�̓��͂���������Œu�������܂��B
'�u������������[�U�[����󂯕t����̂ŁAstrNew�ϐ��Ƀ��[�U�[����̓��͒l���i�[���܂��B

WScript.Echo "�u�����������͂��Ă��������B"
strNew = WScript.StdIn.ReadLine

WScript.Echo "�u" & strNew & "�v�����͂���܂����B�B" 

WScript.Echo "���̕����� : " & strOrg & vbCrLf &_
       "���������� : " & strKey & vbCrLf &_
       "�u�������� : " & strNew & vbCrLf &_
       "�u���̌��� : " &Replace(strOrg,strKey,strNew)
       
