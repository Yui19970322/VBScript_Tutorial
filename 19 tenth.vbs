Option Explicit

Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount
Dim Point

'�|�C���g�̏����l��100�|�C���g
Point = 100

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

        '�ŏI�I�ɐ��������Ƃ��ɁAPoint��0��菬����������A
        'Point��0�ɕύX����B



        MsgBox "����!" & vbCrLf _ 
             & "������ " & Answer & " �ł�!" & vbCrLf _
             & " " & ChallengeCount & " ��ڂɐ������܂���!" & vbCrLf _
             & " " & Point & " Point!!!"
        
        '����1��ڂɐ���������p�[�t�F�N�g�I�ƕ\��
        If ChallengeCount = 1 Then 
            MsgBox "�p�[�t�F�N�g!"
        End If

        Result = True
    ElseIf Abs(Diff) = 1 Then
        'Point��������ƌ��炷
        Point = Minus(Point, 2)
        MsgBox "�߂�!!"
    ElseIf Abs(Diff) < 6 Then
        'Point���������炷
        Point = Minus(Point, 5)
        MsgBox "�ɂ���!"
    ElseIf Abs(Diff) < 11 Then
        'Point�����炷
        Point = Minus(Point, 7)
        MsgBox "��������!"
    ElseIf Abs(Diff) < 16 Then
        'Point���������񌸂炷
        Point = Minus(Point, 9)
        MsgBox "�܂��܂�!"
    ElseIf Abs(Diff) < 19 Then
        'Point�������ƌ��炷
        Point = Minus(Point, 12)
        MsgBox "����!"
    Else
        'Point��傫�����炷
        Point = Minus(Point, 60)
        MsgBox "����!�y���ɉ���!"
    End If
 Wend

'��������Num1�Ƒ�������Num2�𑫂������ʂ�Ԃ��܂��B
Function Add(Num1, Num2)
    'Num1��Num2�𑫂��B���̌��ʂ�Add�ɑ��(�߂�l�Ƃ��ĕԂ��B)
    Add = CInt(Num1 + Num2)
End Function



'��������Num1�����������Num2�����������ʂ�Ԃ��܂��B
Function Minus(Num1, Num2)
    'Num1����Num2�������B���̌��ʂ�Minus�ɑ��(�߂�l�Ƃ��ĕԂ��B)
    Minus = CInt(Num1 - Num2)
End Function



'�����_���Ȑ����𐶐����ĕԂ��܂��B
Function GenerateRandomNumber(Max)
    Randomize
    GenerateRandomNumber = CInt(Rnd * Max)
End Function
