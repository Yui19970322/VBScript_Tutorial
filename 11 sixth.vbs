Option Explicit

Dim Num1
Dim Num2
Dim Number
Dim Answer
Dim Result
Result = False
Num1 = 28
Num2 = 6
Answer = CStr(Num1 * Num2)

While Result = False
    Number = InputBox(Num1 & " ‚©‚¯‚é " & Num2 & " ‚ÍH")
    If Number = Answer Then
        MsgBox "³‰ğ‚Å‚·!"
        Result = True
    End IF
 Wend