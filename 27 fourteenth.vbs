Option Explicit

Dim Number
Number = 0


While Number >= 0
    WScript.Echo "�������l����͂��Ă��������B"
    Number = WScript.StdIn.ReadLine
    WScript.Echo Number & "�����͂���܂����B"
    WScript.Echo "���͂��ꂽ���l��]�艉�Z����B(���͂��ꂽ���l) Mod 2 = " & CInt(Number) Mod 2


Wend

 
'WScript.Echo Number & "�͋����ł��B"
'WScript.Echo Number & "�͊�ł��B"
