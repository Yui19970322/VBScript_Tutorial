Option Explicit

'コマンドプロンプトで実行したときだけ
'Hello! コマンドプロンプトと表示させる
'WScript.echo "Hello! コマンドプロンプト"


Dim strWSH

strWSH = UCase(Right(WScript.FullName,11))

Select Case strWSH
       
       Case "WSCRIPT.EXE"
            MsgBox "GUIで実行されました!" _
                 , vbOkOnly + vbInformation _
                 , "実行モード"

       Case "CSCRIPT.EXE"
            WScript.Echo "Hello! コマンドプロンプト"

End Select