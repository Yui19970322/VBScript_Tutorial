Option Explicit

Dim Number


Number = 10


WScript.Echo Number & " は3の倍数ですか?"
If IsMulOf3(Number) = True Then
    WScript.Echo "3の倍数です。"
Else
    WScript.Echo "3の倍数ではないです。"
End If


WScript.Echo Number & " は5の倍数ですか?"
If IsMulOf5(Number) = True Then
    WScript.Echo "5の倍数です。"
Else
    WScript.Echo "5の倍数ではないです。"
End If


WScript.Echo Number & " は15の倍数ですか?"
If IsMulOf15(Number) = True Then
    WScript.Echo "15の倍数です。"
Else
    WScript.Echo "15の倍数ではないです。"
End If



'3の倍数であれば True を返します。
'3の倍数でない時には False を返します。
Function IsMulOf3(Num1)
    IsMulOf3 = (Num1 Mod 3 = 0)
End Function

'5の倍数であれば True を返します。
'5の倍数でない時には False を返します。
Function IsMulOf5(Num1)
    IsMulOf5 = (Num1 Mod 5 = 0)
End Function

'15の倍数であれば True を返します。
'15の倍数でない時には False を返します。
Function IsMulOf15(Num1)
    IsMulOf15 = (Num1 Mod 15 = 0)
End Function