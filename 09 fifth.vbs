Option Explicit

Dim Num1
Dim Num2
Dim Number
Dim Answer
Num1 = 15
Num2 = 3
Answer = Num1 * Num2

Number = InputBox(Num1 & " かける " & Num2 & " は？")

MsgBox Number & " が入力されました。"
MsgBox "答えは " & Answer & " です。"

