Option Explicit

Dim strWSH

strWSH = UCase(Right(WScript.FullName,11))

Dim Number


Select Case strWSH
       
       Case "CSCRIPT.EXE"
            WScript.Echo "1-20までの数字を入力してください。"
            Number = WScript.StdIn.ReadLine
            
            WScript.Echo Number & " が入力されました!"

End Select