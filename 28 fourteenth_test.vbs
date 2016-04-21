Option Explicit

Dim Number
Number = 0

Dim Result



While Number >= 0
    WScript.Echo "何か数値を入力してください。"
    Number = WScript.StdIn.ReadLine
    WScript.Echo Number & "が入力されました。"
    WScript.Echo "入力された数値を余剰演算する。(入力された数値) Mod 2 = " & CInt(Number) Mod 2

    '入力された数値が偶数だったら「XXは偶数です。」と表示する。
    '入力された数値が奇数だったら「XXは奇数です。」と表示する。
    Result = CInt(Number) Mod 2 

    If Result = 0 Then
        WScript.Echo Number & "は偶数です。"
    ElseIf Result = 1  Then
        WScript.Echo Number & "は奇数です。"
    End IF
Wend

 

