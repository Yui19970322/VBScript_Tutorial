Option Explicit

Dim strWSH

strWSH = UCase(Right(WScript.FullName,11))

Dim Number


Select Case strWSH
       
       Case "CSCRIPT.EXE"
            WScript.Echo "1-20�܂ł̐�������͂��Ă��������B"
            Number = WScript.StdIn.ReadLine
            
            WScript.Echo Number & " �����͂���܂���!"

End Select