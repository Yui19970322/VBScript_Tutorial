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
    Number = InputBox(Num1 & " ������ " & Num2 & " �́H")
    If Number = Answer Then
        MsgBox "�����ł�!"
        Result = True
    End IF
 Wend