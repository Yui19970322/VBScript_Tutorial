Option Explicit

Dim i

'1����50�܂ł̐�����\������B
For i = 1 To 50
    '3�̔{���̂Ƃ��ɂ� Fizz �ƕ\������B
    '5�̔{���̂Ƃ��ɂ� Buzz �ƕ\������B
    '15�̔{���̂Ƃ��ɂ� FizzBuzz �ƕ\������B
    '����ȊO�̎���  -------  �ƕ\������B
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