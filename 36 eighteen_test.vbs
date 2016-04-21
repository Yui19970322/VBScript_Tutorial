Option Explicit

Dim strOrg
Dim strKey
Dim strNew

strOrg = "ここで履物を脱いでください"
strKey = "履物"

'ユーザーからの入力を受け付けます。
'"ここで履物を脱いでください" "履物" を、ユーザーの入力した文字列で置き換えます。
'置換文字列をユーザーから受け付けるので、strNew変数にユーザーからの入力値を格納します。

WScript.Echo "置換文字列を入力してください。"
strNew = WScript.StdIn.ReadLine

WScript.Echo "「" & strNew & "」が入力されました。。" 

WScript.Echo "元の文字列 : " & strOrg & vbCrLf &_
       "検索文字列 : " & strKey & vbCrLf &_
       "置換文字列 : " & strNew & vbCrLf &_
       "置換の結果 : " &Replace(strOrg,strKey,strNew)
       
