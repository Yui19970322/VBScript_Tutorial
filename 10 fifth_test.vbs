Option Explicit

Dim Num1
Dim Num2
Dim Number
Dim Answer
Num1 = 28
Num2 = 6
Answer = CStr(Num1 * Num2)

Number = InputBox(Num1 & " ������ " & Num2 & " �́H")

MsgBox Number & " �����͂���܂����B"
'MsgBox "������ " & Answer & " �ł��B"

'���������[�U�[�̓���������(Number = Answer)�Ȃ�u�����ł�!�v�ƕ\������B
If Number = Answer Then
    MsgBox "�����ł�!"
'�s�����������͖��񓚂̏ꍇ�́u�͂���ł�!�v�ƕ\������B
Else
    MsgBox "�͂���ł�!"
End IF    
 