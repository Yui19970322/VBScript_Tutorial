Option Explicit

Dim i

'1から50までの数字を表示する。
For i = 1 To 50
    '3の倍数のときには Fizz と表示する。
    '5の倍数のときには Buzz と表示する。
    '15の倍数のときには FizzBuzz と表示する。
    'それ以外の時は  -------  と表示する。
    ' ex)
    '1 ------- 
    '2 ------- 
    '3 Fizz 
    '4 ------- 
    '5 Buzz 
    '6 Fizz 
    '7 ------- 
    '8 ------- 
    '9 Fizz 
    '10 Buzz 
    '11 ------- 
    '12 ------- 
    '13 ------- 
    '14 ------- 
    '15 FizzBuzz 
    '16 ------- 
    '......
    

If IsMulOf15(i) = True Then
'WScript.StdOut.WriteLine i & ""
	WScript.StdOut.WriteLine i & " FizzBuzz "

ElseIf IsMulOf5(i) = True Then
	WScript.StdOut.WriteLine i & " Buzz "

ElseIf IsMulOf3(i) = True Then
	WScript.StdOut.WriteLine i & " Fizz "

Else
	WScript.StdOut.WriteLine i & " ------- "

End If
Next


'WScript.StdOut.WriteLine i & " ------- "


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