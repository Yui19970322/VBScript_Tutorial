Option Explicit

WScript.Echo "WScriptで取得できる情報"
WScript.Echo WScript.ScriptName

WScript.Echo "--------------------------------------"


WScript.Echo "WScript.ScriptFullNameで取得できる情報"
WScript.Echo WScript.ScriptFullName


WScript.Echo "--------------------------------------"


WScript.Echo "スクリプトが保存されている場所"
WScript.Echo Replace(WScript.ScriptFullName, WScript.ScriptName, "")