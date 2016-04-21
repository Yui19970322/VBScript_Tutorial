Option Explicit

Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount

'答えを入力した数をカウント
ChallengeCount = 0
Result = False

Randomize
Answer = CInt(Rnd * 20)


While Result = False
    Number = InputBox("コンピューターが生成した数を入力してください。" & vbCrLf & "(答えは " & Answer & " です)")
    '答えを入力した数をカウント
    ChallengeCount = ChallengeCount + 1
    
    Diff = Answer - CInt(Number)
    If Diff = 0 Then
        MsgBox "正解!" & vbCrLf & "答えは " & Answer & " です!" & vbCrLf & " " & ChallengeCount & " 回目に正解しました!"
        
        'もし1回目に正解したらパーフェクト！と表示
        If ChallengeCount = 1 Then 
            MsgBox "パーフェクト!"
        End If
        
        
        
        Result = True
    ElseIf Abs(Diff) = 1 Then
        MsgBox "近い!!"
    ElseIf Abs(Diff) < 6 Then
        MsgBox "惜しい!"
    ElseIf Abs(Diff) < 11 Then
        MsgBox "もう少し!"
    ElseIf Abs(Diff) < 16 Then
        MsgBox "まだまだ!"
    ElseIf Abs(Diff) < 19 Then
        MsgBox "遠い!"
    Else
        MsgBox "遠い!遥かに遠い!"
    End If
 Wend