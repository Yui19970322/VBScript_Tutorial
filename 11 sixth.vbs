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
    Number = InputBox(Num1 & " かける " & Num2 & " は？")
    If Number = Answer Then
        MsgBox "正解です!"
        Result = True
    End IF
 Wend