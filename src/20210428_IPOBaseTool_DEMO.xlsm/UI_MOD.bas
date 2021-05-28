Attribute VB_Name = "UI_MOD"
Option Explicit
'��UI�\������n���W���[��

Public Sub OP_form0()
'���Ǘ��\ID���̓t�H�[���N��
    Dim eRow As Long
    Dim Ans As String
    
    eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
    ActiveSheet.Unprotect
    If Cells(eRow, 7).Value = "" Then
        Cells(eRow, 7).Select
        Ans = MsgBox("�O���J����ID���ݒ肳��Ă��Ȃ��Ǘ��\ID������܂�" & vbCrLf & _
        "�܂��ݒ肳��Ă��Ȃ��O���J����ID��ݒ肵�Ă�������" & vbCrLf & _
        "���ݒ�̊Ǘ��\ID��j�����đ����܂����H", vbYesNo + vbInformation, "�O���J����ID�����ݒ��ID������܂�")
        If Ans = vbNo Then
            Call St_Lock
            End
        ElseIf Ans = vbYes Then
            Cells(eRow, 5).ClearContents
        End If
    End If
    Call St_Lock
    UF_0.Show

End Sub

Public Sub OP_form1()
'���O��ID���̓t�H�[���N��
    Dim Ws As Worksheet
    Dim eRow As Long
    
    Set Ws = Sheets("�J�����ݒ�")
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "�Ǘ��\�J����ID��ݒ肵�Ă���s���Ă�������", 16, "�Ǘ��\�J����ID�����̓G���["
            Exit Sub
        End If
    End With
    UF_1.Show

End Sub

Public Sub ALL_Reset()
'��UI���Z�b�g
    Call St_AllUnvis
    With Sheets("�C���|�[�g")
        .Unprotect
        .Shapes("Fil_1").Visible = False
        .Shapes("Gr_1").Visible = False
    End With

End Sub

Public Sub St_AllUnvis()
'���z�[���V�[�g�ȊO�̃V�[�g��\��
    Dim Stc As Long
    Dim i As Long
    
    Stc = ThisWorkbook.Sheets.Count
    Sheets("�z�[��").Visible = True
    For i = 2 To Stc
        Sheets(i).Visible = False
    Next i

End Sub

Public Sub vis_UISt()
'���z�[����ʃV�[�g�\��
    Call vis_St("�z�[��")
    
End Sub

Public Sub vis_ImportSt()
'���C���|�[�g��ʃV�[�g�\��
    Call vis_St("�C���|�[�g")
    ActiveSheet.Unprotect
    ActiveSheet.Range("C7").ClearContents
    Call St_Lock

End Sub

Public Sub vis_CSVImportSt()
'���C���|�[�g��ʃV�[�g�\��
    Call vis_St("CSV�C���|�[�g")
    ActiveSheet.Unprotect
    ActiveSheet.Range("C7").ClearContents
    Call St_Lock

End Sub

Public Sub vis_KANRISt()
'���Ǘ��\�ҏW�o�^�V�[�g�\��
    Application.ScreenUpdating = False
    Call vis_St("�Ǘ��\�ҏW�o�^")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call Re_KANRI
    Call Re_Scrl
    ActiveSheet.Range("E10").Select

End Sub

Public Sub vis_RegNewIDSt()
'���Ǘ��\�V�K�o�^�V�[�g�\��
    Call vis_St("�Ǘ��\�V�K�o�^")
    Sheets("�Ǘ��\�V�K�o�^").Unprotect
    Sheets("�Ǘ��\�V�K�o�^").Range("D6").ClearContents
    Call St_Lock
    
End Sub

Public Sub vis_CosKANRISt()
'���J�X�^���ҏW�o�^�V�[�g�\��
    Application.ScreenUpdating = False
    Call vis_St("�Ǘ��\�ҏW�o�^")
    Call Re_CosKANRI
    Call Re_Scrl
    
End Sub

Public Sub vis_KANRIvewSt()
'���Ǘ��\�o�̓r���[�V�[�g�\��
    Application.ScreenUpdating = False
    Call vis_St("�Ǘ��\�o�̓r���[")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call Run_Douki_KANRIvew
    Call Re_Scrl
    Call St_Lock

End Sub

Public Sub vis_CostumvewSt()
'���J�X�^���r���[�V�[�g�\��
    Application.ScreenUpdating = False
    Call vis_St("�J�X�^���r���[")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call St_Lock
'    Call Run_Search_Costumvew
  
End Sub

Public Sub vis_T_GAIBSt()
'���O���f�[�^�V�[�g�\��
    Call vis_St("�O���f�[�^")
    Call Run_Douki_GAIB
    Call Re_Scrl
    Call Clear_SearchKey_G
    
End Sub

Public Sub vis_SETTEISt()
'���ݒ�V�[�g�\��
    Call vis_St("�ݒ�")

End Sub

Public Sub vis_DBSETTEISt()
'���f�[�^�x�[�X�ݒ�V�[�g�\��
    Call vis_St("�f�[�^�x�[�X�ݒ�")

End Sub
Public Sub vis_SETDirectSt()
'���f�B���N�g���ݒ�V�[�g�\��
    Call vis_St("�f�B���N�g���ݒ�")

End Sub

Public Sub vis_SETFildsSt()
'���Ǘ��\�t�B�[���h�ݒ�V�[�g�\��
    Call vis_St("�Ǘ��\�t�B�[���h�ݒ�")

End Sub

Public Sub vis_RAKRangSt()
'���V�[�g�͈͐ݒ�V�[�g�\��
    Call vis_St("�O���f�[�^�V�[�g�͈͐ݒ�")

End Sub

Public Sub vis_TOESET1St()
'���J�����ݒ�1�V�[�g�\��
    Call vis_St("�J�����ݒ�")
    Call Re_Scrl
    
End Sub

Public Function vis_St(ByVal Op_St As String)
'���w��V�[�g�̕\���ˊJ���Ă���V�[�g�̔�\��
    With Sheets(Op_St)
        .Visible = True
        ActiveSheet.Visible = False
    End With
    Call St_Lock
    
End Function

Sub St_Lock()
'���V�[�g���b�N
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub


Public Sub Re_Scrl()
'���X�N���[�����Z�b�g
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
End Sub
