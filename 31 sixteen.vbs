Option Explicit

Dim objWShell
Set objWShell = Createobject("WScript.Shell")

'ユーザーから数値の入力を受け付けます。
'ユーザーの入力した数値が偶数か奇数かを判定します。
'ユーザーの入力した数値が偶数の時は、Popupを3秒間表示します。
'ユーザーの入力した数値が奇数の時は、Popupをユーザーの入力した数値の秒数だけ表示します。





objWShell.Popup "おはようございます" _ 
              , 5 _
              , "Popupサンプル" _
              ,vbOkOnly + vbInformation


Set objWShell = Nothing
