Option Explicit

Dim Number


Number = 10


WScript.Echo Number & " ‚Í3‚Ì”{”‚Å‚·‚©?"
If IsMulOf3(Number) = True Then
    WScript.Echo "3‚Ì”{”‚Å‚·B"
Else
    WScript.Echo "3‚Ì”{”‚Å‚Í‚È‚¢‚Å‚·B"
End If


WScript.Echo Number & " ‚Í5‚Ì”{”‚Å‚·‚©?"
If IsMulOf5(Number) = True Then
    WScript.Echo "5‚Ì”{”‚Å‚·B"
Else
    WScript.Echo "5‚Ì”{”‚Å‚Í‚È‚¢‚Å‚·B"
End If


WScript.Echo Number & " ‚Í15‚Ì”{”‚Å‚·‚©?"
If IsMulOf15(Number) = True Then
    WScript.Echo "15‚Ì”{”‚Å‚·B"
Else
    WScript.Echo "15‚Ì”{”‚Å‚Í‚È‚¢‚Å‚·B"
End If



'3‚Ì”{”‚Å‚ ‚ê‚Î True ‚ğ•Ô‚µ‚Ü‚·B
'3‚Ì”{”‚Å‚È‚¢‚É‚Í False ‚ğ•Ô‚µ‚Ü‚·B
Function IsMulOf3(Num1)
    IsMulOf3 = (Num1 Mod 3 = 0)
End Function

'5‚Ì”{”‚Å‚ ‚ê‚Î True ‚ğ•Ô‚µ‚Ü‚·B
'5‚Ì”{”‚Å‚È‚¢‚É‚Í False ‚ğ•Ô‚µ‚Ü‚·B
Function IsMulOf5(Num1)
    IsMulOf5 = (Num1 Mod 5 = 0)
End Function

'15‚Ì”{”‚Å‚ ‚ê‚Î True ‚ğ•Ô‚µ‚Ü‚·B
'15‚Ì”{”‚Å‚È‚¢‚É‚Í False ‚ğ•Ô‚µ‚Ü‚·B
Function IsMulOf15(Num1)
    IsMulOf15 = (Num1 Mod 15 = 0)
End Function