Option Explicit

Dim strOrg
Dim strKey
Dim strNew

strOrg = "ここで履物を脱いでください"
strKey = "履物"
strNew = "は着物"
MsgBox "元の文字列 : " & strOrg & vbCr &_
       "検索文字列 : " & strKey & vbCr &_
       "置換文字列 : " & strNew & vbCr &_
       "置換の結果 : " &Replace(strOrg,strKey,strNew)
        


