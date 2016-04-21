Option Explicit

Dim strWSH

strWSH = UCase(Right(WScript.FullName,11))

Dim Number


Select Case strWSH
       
       Case "CSCRIPT.EXE"
            WScript.Echo "1-20‚Ü‚Å‚Ì”š‚ğ“ü—Í‚µ‚Ä‚­‚¾‚³‚¢B"
            Number = WScript.StdIn.ReadLine
            
            WScript.Echo Number & " ‚ª“ü—Í‚³‚ê‚Ü‚µ‚½!"

End Select