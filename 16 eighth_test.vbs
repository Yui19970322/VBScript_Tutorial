Option Explicit

Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount

'��������͂��������J�E���g
ChallengeCount = 0
Result = False

Answer = GenerateRandomNumber(20)


While Result = False
    Number = InputBox("�R���s���[�^�[����������������͂��Ă��������B" & vbCrLf & "(������ " & Answer & " �ł�)")

    'Add�֐����g���� ChallengeCount �𖈉�X�V����B
    ChallengeCount = Add(ChallengeCount, 1)


    
    Diff = Answer - CInt(Number)
    If Diff = 0 Then
        MsgBox "����!" & vbCrLf & "������ " & Answer & " �ł�!" & vbCrLf & " " & ChallengeCount & " ��ڂɐ������܂���!"
        
        '����1��ڂɐ���������p�[�t�F�N�g�I�ƕ\��
        If ChallengeCount = 1 Then 
            MsgBox "�p�[�t�F�N�g!"
        End If

        Result = True
    ElseIf Abs(Diff) = 1 Then
        MsgBox "�߂�!!"
    ElseIf Abs(Diff) < 6 Then
        MsgBox "�ɂ���!"
    ElseIf Abs(Diff) < 11 Then
        MsgBox "��������!"
    ElseIf Abs(Diff) < 16 Then
        MsgBox "�܂��܂�!"
    ElseIf Abs(Diff) < 19 Then
        MsgBox "����!"
    Else
        MsgBox "����!�y���ɉ���!"
    End If
 Wend

'��������Num1�Ƒ�������Num2�𑫂������ʂ�Ԃ��܂��B
Function Add(Num1, Num2)
    'Num1��Num2�𑫂��B���̌��ʂ�Add�ɑ��(�߂�l�Ƃ��ĕԂ��B)
    Add = CInt(Num1 + Num2)
End Function

'�����_���Ȑ����𐶐����ĕԂ��܂��B
Function GenerateRandomNumber(Max)
    Randomize
    GenerateRandomNumber = CInt(Rnd * Max)
End Function
