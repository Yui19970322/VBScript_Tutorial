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
        
        Result = True
    ElseIf Abs(Diff) = 1 Then
        MsgBox "?????!"
    ElseIf Abs(Diff) < 6 Then
        MsgBox "???????!"
    ElseIf Abs(Diff) < 11 Then
        MsgBox "????????????!"
    ElseIf Abs(Diff) < 16 Then
        MsgBox "??????!"
    ElseIf Abs(Diff) < 19 Then
        MsgBox "??????!??????!"
    Else
        MsgBox "?????I?y????????I"
    End If
 Wend