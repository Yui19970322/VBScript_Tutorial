Option Explicit

Dim Number
Number = 0

Dim Result



While Number >= 0
    WScript.Echo "�������l����͂��Ă��������B"
    Number = WScript.StdIn.ReadLine
    WScript.Echo Number & "�����͂���܂����B"
    WScript.Echo "���͂��ꂽ���l��]�艉�Z����B(���͂��ꂽ���l) Mod 2 = " & CInt(Number) Mod 2

    '���͂��ꂽ���l��������������uXX�͋����ł��B�v�ƕ\������B
    '���͂��ꂽ���l�����������uXX�͊�ł��B�v�ƕ\������B
    Result = CInt(Number) Mod 2 

    If Result = 0 Then
        WScript.Echo Number & "�͋����ł��B"
    ElseIf Result = 1  Then
        WScript.Echo Number & "�͊�ł��B"
    End IF
Wend

 

