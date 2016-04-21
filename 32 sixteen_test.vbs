Option Explicit

Dim Number
Dim objWShell
Dim Result
Set objWShell = Createobject("WScript.Shell")


'ユーザーから数値の入力を受け付けます。
WScript.Echo "何か数値を入力してください。"
Number = WScript.StdIn.ReadLine
WScript.Echo Number & "が入力されました。"

'ユーザーの入力した数値が偶数か奇数かを判定します。
Result = CInt(Number) Mod 2
If Result = 0 Then
    WScript.Echo Number & "偶数です"
ElseIf Result = 1  Then
    WScript.Echo Number & "奇数です"
End IF


'ユーザーの入力した数値が偶数の時は、Popupを3秒間表示します。
If Result = 0 Then
    objWShell.Popup "こんにちは" _ 
              , 3 _
              , "Popupサンプル" _
              ,vbOkOnly + vbInformation
End If             

'ユーザーの入力した数値が奇数の時は、Popupをユーザーの入力した数値の秒数だけ表示します。
If Result = 1 Then
    objWShell.Popup "こんばんは" _ 
              ,(Number) _ 
              , "Popupサンプル" _
              ,vbOkOnly + vbInformation
End If             

Set objWShell = Nothing
