Attribute VB_Name = "ModDeleteRowArray"
Option Explicit

'DeleteRowArray    �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2D      �E�E�E���ꏊ�FFukamiAddins3.ModArray
'CheckArray2DStart1�E�E�E���ꏊ�FFukamiAddins3.ModArray

'------------------------------


'�z��̏����֌W�̃v���V�[�W��

'------------------------------


Public Function DeleteRowArray(Array2D, DeleteRow&)
'�񎟌��z��̎w��s�����������z����o�͂���
'20210917

'����
'Array2D  �E�E�E�񎟌��z��
'DeleteRow�E�E�E��������s�ԍ�

    '�����`�F�b�N
    Call CheckArray2D(Array2D, "Array2D")
    Call CheckArray2DStart1(Array2D, "Array2D")
    
    Dim I&, J&, K&, M&, N& '�����グ�p(Long�^)
    N = UBound(Array2D, 1) '�s��
    M = UBound(Array2D, 2) '��
    
    If DeleteRow < 1 Then
        MsgBox ("�폜����s�ԍ���1�ȏ�̒l�����Ă�������")
        Stop
        End
    ElseIf DeleteRow > N Then
        MsgBox ("�폜����s�ԍ��͌��̓񎟌��z��̍s��" & N & "�ȉ��̒l�����Ă�������")
        Stop
        End
    End If
    
    '����
    Dim Output
    ReDim Output(1 To N - 1, 1 To M)
    K = 0
    For I = 1 To N
        If I <> DeleteRow Then
            K = K + 1
            For J = 1 To M
                Output(K, J) = Array2D(I, J)
            Next J
        End If
    Next I
    
    '�o��
    DeleteRowArray = Output

End Function

Private Sub CheckArray2D(InputArray, Optional HairetuName$ = "�z��")
'���͔z��2�����z�񂩂ǂ����`�F�b�N����
'20210804

    Dim Dummy2%, Dummy3%
    On Error Resume Next
    Dummy2 = UBound(InputArray, 2)
    Dummy3 = UBound(InputArray, 3)
    On Error GoTo 0
    If Dummy2 = 0 Or Dummy3 <> 0 Then
        MsgBox (HairetuName & "��2�����z�����͂��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub

Private Sub CheckArray2DStart1(InputArray, Optional HairetuName$ = "�z��")
'����2�����z��̊J�n�ԍ���1���ǂ����`�F�b�N����
'20210804

    If LBound(InputArray, 1) <> 1 Or LBound(InputArray, 2) <> 1 Then
        MsgBox (HairetuName & "�̊J�n�v�f�ԍ���1�ɂ��Ă�������")
        Stop
        Exit Sub '���͌��̃v���V�[�W�����m�F���邽�߂ɔ�����
    End If

End Sub


