Option Explicit

Dim Number
Number = 0


While Number >= 0
    WScript.Echo "何か数値を入力してください。"
    Number = WScript.StdIn.ReadLine
    WScript.Echo Number & "が入力されました。"
    WScript.Echo "入力された数値を余剰演算する。(入力された数値) Mod 2 = " & CInt(Number) Mod 2


Wend

 
'WScript.Echo Number & "は偶数です。"
'WScript.Echo Number & "は奇数です。"
