Option Explicit

WScript.Echo "WScript�Ŏ擾�ł�����"
WScript.Echo WScript.ScriptName

WScript.Echo "--------------------------------------"


WScript.Echo "WScript.ScriptFullName�Ŏ擾�ł�����"
WScript.Echo WScript.ScriptFullName


WScript.Echo "--------------------------------------"


WScript.Echo "�X�N���v�g���ۑ�����Ă���ꏊ"
WScript.Echo Replace(WScript.ScriptFullName, WScript.ScriptName, "")