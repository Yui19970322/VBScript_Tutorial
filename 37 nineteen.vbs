Option Explicit

Dim Number


Number = 10


WScript.Echo Number & " ��3�̔{���ł���?"
If IsMulOf3(Number) = True Then
    WScript.Echo "3�̔{���ł��B"
Else
    WScript.Echo "3�̔{���ł͂Ȃ��ł��B"
End If


WScript.Echo Number & " ��5�̔{���ł���?"
If IsMulOf5(Number) = True Then
    WScript.Echo "5�̔{���ł��B"
Else
    WScript.Echo "5�̔{���ł͂Ȃ��ł��B"
End If


WScript.Echo Number & " ��15�̔{���ł���?"
If IsMulOf15(Number) = True Then
    WScript.Echo "15�̔{���ł��B"
Else
    WScript.Echo "15�̔{���ł͂Ȃ��ł��B"
End If



'3�̔{���ł���� True ��Ԃ��܂��B
'3�̔{���łȂ����ɂ� False ��Ԃ��܂��B
Function IsMulOf3(Num1)
    IsMulOf3 = (Num1 Mod 3 = 0)
End Function

'5�̔{���ł���� True ��Ԃ��܂��B
'5�̔{���łȂ����ɂ� False ��Ԃ��܂��B
Function IsMulOf5(Num1)
    IsMulOf5 = (Num1 Mod 5 = 0)
End Function

'15�̔{���ł���� True ��Ԃ��܂��B
'15�̔{���łȂ����ɂ� False ��Ԃ��܂��B
Function IsMulOf15(Num1)
    IsMulOf15 = (Num1 Mod 15 = 0)
End Function