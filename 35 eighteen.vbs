Option Explicit

Dim strOrg
Dim strKey
Dim strNew

strOrg = "�����ŗ�����E���ł�������"
strKey = "����"
strNew = "�͒���"
MsgBox "���̕����� : " & strOrg & vbCr &_
       "���������� : " & strKey & vbCr &_
       "�u�������� : " & strNew & vbCr &_
       "�u���̌��� : " &Replace(strOrg,strKey,strNew)
        


