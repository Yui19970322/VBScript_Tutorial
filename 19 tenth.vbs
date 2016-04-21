Option Explicit

Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount
Dim Point

'ポイントの初期値は100ポイント
Point = 100

'答えを入力した数をカウント
ChallengeCount = 0
Result = False

Answer = GenerateRandomNumber(20)


While Result = False
    Number = InputBox("コンピューターが生成した数を入力してください。" & vbCrLf & "(答えは " & Answer & " です)")

    'Add関数を使って ChallengeCount を毎回更新する。
    ChallengeCount = Add(ChallengeCount, 1)

    Diff = Answer - CInt(Number)
    If Diff = 0 Then

        '最終的に正解したときに、Pointが0より小さかったら、
        'Pointを0に変更する。



        MsgBox "正解!" & vbCrLf _ 
             & "答えは " & Answer & " です!" & vbCrLf _
             & " " & ChallengeCount & " 回目に正解しました!" & vbCrLf _
             & " " & Point & " Point!!!"
        
        'もし1回目に正解したらパーフェクト！と表示
        If ChallengeCount = 1 Then 
            MsgBox "パーフェクト!"
        End If

        Result = True
    ElseIf Abs(Diff) = 1 Then
        'Pointをちょっと減らす
        Point = Minus(Point, 2)
        MsgBox "近い!!"
    ElseIf Abs(Diff) < 6 Then
        'Pointを少し減らす
        Point = Minus(Point, 5)
        MsgBox "惜しい!"
    ElseIf Abs(Diff) < 11 Then
        'Pointを減らす
        Point = Minus(Point, 7)
        MsgBox "もう少し!"
    ElseIf Abs(Diff) < 16 Then
        'Pointをたくさん減らす
        Point = Minus(Point, 9)
        MsgBox "まだまだ!"
    ElseIf Abs(Diff) < 19 Then
        'Pointをもっと減らす
        Point = Minus(Point, 12)
        MsgBox "遠い!"
    Else
        'Pointを大きく減らす
        Point = Minus(Point, 60)
        MsgBox "遠い!遥かに遠い!"
    End If
 Wend

'第一引数のNum1と第二引数のNum2を足した結果を返します。
Function Add(Num1, Num2)
    'Num1とNum2を足す。その結果をAddに代入(戻り値として返す。)
    Add = CInt(Num1 + Num2)
End Function



'第一引数のNum1から第二引数のNum2を引いた結果を返します。
Function Minus(Num1, Num2)
    'Num1からNum2を引く。その結果をMinusに代入(戻り値として返す。)
    Minus = CInt(Num1 - Num2)
End Function



'ランダムな数字を生成して返します。
Function GenerateRandomNumber(Max)
    Randomize
    GenerateRandomNumber = CInt(Rnd * Max)
End Function
