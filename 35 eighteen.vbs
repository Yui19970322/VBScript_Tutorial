Option Explicit

Dim strOrg
Dim strKey
Dim strNew

strOrg = "‚±‚±‚Å—š•¨‚ğ’E‚¢‚Å‚­‚¾‚³‚¢"
strKey = "—š•¨"
strNew = "‚Í’…•¨"
MsgBox "Œ³‚Ì•¶š—ñ : " & strOrg & vbCr &_
       "ŒŸõ•¶š—ñ : " & strKey & vbCr &_
       "’uŠ·•¶š—ñ : " & strNew & vbCr &_
       "’uŠ·‚ÌŒ‹‰Ê : " &Replace(strOrg,strKey,strNew)
        


