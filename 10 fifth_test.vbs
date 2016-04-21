Option Explicit

Dim Num1
Dim Num2
Dim Number
Dim Answer
Num1 = 28
Num2 = 6
Answer = CStr(Num1 * Num2)

Number = InputBox(Num1 & " かける " & Num2 & " は？")

MsgBox Number & " が入力されました。"
'MsgBox "答えは " & Answer & " です。"

'もしもユーザーの答えが正解(Number = Answer)なら「正解です!」と表示する。
If Number = Answer Then
    MsgBox "正解です!"
'不正解もしくは未回答の場合は「はずれです!」と表示する。
Else
    MsgBox "はずれです!"
End IF    
 