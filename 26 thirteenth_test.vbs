Option Explicit

'�����ăQ�[�����R�}���h�v�����v�g�Ŏ��s�����Ƃ��ɂ́A�R�}���h�v�����v�g�����̕\���Ŏ��s�ł���悤�ɂ���B
'�R�}���h�v�����v�g�Ŏ��s�����Ƃ��ɂ̓��b�Z�[�W�{�b�N�X�͏o�Ȃ��悤�ɂ���B
'�����ăQ�[����GUI�Ŏ��s�����Ƃ��ɂ́AGUI�����̕\���Ŏ��s�ł���悤�ɂ���B

Dim strWSH
Dim Number
Dim Answer
Dim Result
Dim Diff
Dim ChallengeCount
Dim Point

strWSH = UCase(Right(WScript.FullName,11))

'�|�C���g�̏����l��100�|�C���g
Point = 100

'��������͂��������J�E���g
ChallengeCount = 0
Result = False

Answer = GenerateRandomNumber(20)


While Result = False
    Select Case strWSH   
        Case "CSCRIPT.EXE"
            WScript.Echo "1-20�܂ł̐�������͂��Ă��������B"
            Number = WScript.StdIn.ReadLine
        Case "WSCRIPT.EXE"
            Number = InputBox("1-20�܂ł̐�������͂��Ă��������B")
    End Select


    'Add�֐����g���� ChallengeCount �𖈉�X�V����B
    ChallengeCount = Add(ChallengeCount, 1)

    Diff = Answer - CInt(Number)
    If Diff = 0 Then

        '�ŏI�I�ɐ��������Ƃ��ɁAPoint��0��菬����������A
        'Point��0�ɕύX����B
        If Point < 0 Then
            Point = 0
        End If

        WScript.echo "����!" & vbCrLf _ 
             & "������ " & Answer & " �ł�!" & vbCrLf _
             & " " & ChallengeCount & " ��ڂɐ������܂���!" & vbCrLf _
             & " " & Point & " Point!!!"
        
        '����1��ڂɐ���������p�[�t�F�N�g�I�ƕ\��
        If ChallengeCount = 1 Then 
            WScript.echo "�p�[�t�F�N�g!"
        End If

        Result = True
    ElseIf Abs(Diff) = 1 Then
        'Point��������ƌ��炷
        Point = Minus(Point, 2)
        WScript.echo "�߂�!!"
    ElseIf Abs(Diff) < 6 Then
        'Point���������炷
        Point = Minus(Point, 5)
        WScript.echo "�ɂ���!"
    ElseIf Abs(Diff) < 11 Then
        'Point�����炷
        Point = Minus(Point, 7)
        WScript.echo "��������!"
    ElseIf Abs(Diff) < 16 Then
        'Point���������񌸂炷
        Point = Minus(Point, 9)
        WScript.echo "�܂��܂�!"
    ElseIf Abs(Diff) < 19 Then
        'Point�������ƌ��炷
        Point = Minus(Point, 12)
        WScript.echo "����!"
    Else
        'Point��傫�����炷
        Point = Minus(Point, 60)
        WScript.echo "����!�y���ɉ���!"
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





