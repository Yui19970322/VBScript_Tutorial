Option Explicit

Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount

'��������͂��������J�E���g
ChallengeCount = 0
Result = False

Randomize
Answer = CInt(Rnd * 20)


While Result = False
    Number = InputBox("�R���s���[�^�[����������������͂��Ă��������B" & vbCrLf & "(������ " & Answer & " �ł�)")
    '��������͂��������J�E���g
    ChallengeCount = ChallengeCount + 1
    
    Diff = Answer - CInt(Number)
    If Diff = 0 Then
        MsgBox "����!" & vbCrLf & "������ " & Answer & " �ł�!" & vbCrLf & " " & ChallengeCount & " ��ڂɐ������܂���!"
        
        '����1��ڂɐ���������p�[�t�F�N�g�I�ƕ\��
        
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