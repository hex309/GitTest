Attribute VB_Name = "St_MOD"
Option Explicit
'���V�[�g����n���W���[��

Public Sub Run_FeldsSET_Seve()
'���t�B�[���h�ݒ�ۑ����f���s
    Dim Ws As Worksheet
    
    Application.ScreenUpdating = False
    Set Ws = Sheets("�Ǘ��\�t�B�[���h�ݒ�")
    Call Get_KANRIColList '�Ǘ��\�t�B�[���h�����X�g�쐬
    ThisWorkbook.Save
    MsgBox "�ݒ�̕ۑ��E���f���������܂���", vbInformation
    
End Sub

Sub Run_Col_Seve()
'���J�����ݒ�ۑ����f���s
    Call vis_CosKANRISt
'    Call Col_Seve("TG_T_ColList", 0)
'    Call Col_Seve("TG_G_ColList", 1)
'    Call St_ColLis("TG_G_ColList", "�O���f�[�^", "B5")
    
    ThisWorkbook.Save
    
End Sub

Public Function Col_Seve(ByVal str_L_Stn As String, Flg As Long)
'���J�����ݒ�ۑ����f
    '�ݒ�V�[�g����ݒ�J�������̂݃J�������V�[�g�ɓ]�L
    '����1:�����V�[�g��,�ۑ��J�����I���t���O�@�O������/1=�O��
    Dim eRow As Long
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim str_StRng As String
  
    Set R_Ws = Sheets("�J�����ݒ�")
    Set L_Ws = Sheets(str_L_Stn)
 '�Ǐo�f�[�^�Z�b�g *******************************************************************
    If Flg = 0 Then '���ЃJ�����f�[�^�Ǐo
        eRow = R_Ws.Cells(Rows.Count, 5).End(xlUp).Row
        str_StRng = "�J�����ݒ�$E4:E" & eRow
        Call Opn_ExlRs(str_StRng, "�Ǘ��\�J����ID")
    End If
    If Flg = 1 Then  '�O���J�����f�[�^�Ǐo
        eRow = R_Ws.Cells(Rows.Count, 7).End(xlUp).Row
        str_StRng = "�J�����ݒ�$G4:G" & eRow
        Call Opn_ExlRs(str_StRng, "�O���J����ID")
    End If
 '�Ǐo�f�[�^�Z�b�g�����܂� **************************************************************
 '�f�[�^�]�L
    With L_Ws
       .Unprotect
       .Cells.ClearContents
       .Range("A1").CopyFromRecordset Exl_Rs
    End With
    Call Dis_Exl_Rs

End Function

Public Function St_ColLis(ByVal str_R_Stn As String, str_L_Stn As String, str_rng As String)
'���O���f�[�^�V�[�g�ւ̊O���J�����ݒ蔽�f
    '����1:�Ǐo�V�[�g��,����2:�����V�[�g��,����3:�t�B�[���h�擪�Z���A�h���X
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim LC_Ws As Worksheet
    Dim str_LC_Stn As String
    Dim sRow, sCol As Long
    Dim cnt As Long

    Set R_Ws = Sheets(str_R_Stn)
    Set L_Ws = Sheets(str_L_Stn)

    sRow = L_Ws.Range(str_rng).Row
    sCol = L_Ws.Range(str_rng).Column
    
    Call Opn_ExlRs("�J�����ݒ�$G4:G300", "�O���J����ID")

    With L_Ws
        .Unprotect
        .Range(.Cells(sRow, sCol), .Cells(sRow, 300)).ClearContents '�����t�B�[���h���N���A
        cnt = 0
        Do Until Exl_Rs.EOF
            cnt = cnt + 1
            .Cells(sRow, sCol - 1 + cnt).Value = Exl_Rs!�O���J����ID
            Exl_Rs.MoveNext
        Loop
    End With
    L_Ws.Range("G:GZ").EntireColumn.AutoFit
    Call Dis_Exl_Rs

End Function

Public Sub Re_KANRI()
'���Ǘ��\�ҏW_�o�^�V�[�g���Z�b�g
    Dim K_Ws As Worksheet
    
    Set K_Ws = Sheets("�Ǘ��\�ҏW�o�^")
     With K_Ws
        .Unprotect
        .Rows(4).ClearContents
        .Rows(10).ClearContents
        .Range("11:100000").Delete
'        .Range("B10").Select
    End With
    Call Re_Scrl
    Call St_Lock

End Sub

Public Sub Re_CosKANRI()
'���Ǘ��\�ҏW_�o�^�V�[�g���Z�b�g
    Dim K_Ws As Worksheet
    
    Set K_Ws = Sheets("�Ǘ��\�ҏW�o�^")
     With K_Ws
        .Unprotect
        .Rows(4).ClearContents
        .Rows(10).ClearContents
        .Range("11:100000").Delete
'        .Range("B10").Select
    End With
    Call Re_Scrl
    Call St_Lock

End Sub
Public Sub Run_Clear_SearchKey1()
'�����������N���A�@�Ǘ��\�ҏW_�o�^�p
    Dim Ws As Worksheet
    
    Set Ws = Sheets("�Ǘ��\�ҏW�o�^")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Re_KANRI

End Sub

Public Sub Run_Clear_SearchKey2()
'�����������N���A�@�Ǘ��\�o�̓r���[�p
    Dim Ws As Worksheet
    
    Application.ScreenUpdating = False
    Set Ws = Sheets("�Ǘ��\�o�̓r���[")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Run_Douki_KANRIvew

End Sub

Public Sub Run_Clear_SearchKey3()
'�����������N���A�@ �J�X�^���r���[�p
    Dim Ws As Worksheet
    
    Set Ws = Sheets("�J�X�^���r���[")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Run_Search_Costumvew

End Sub

Public Sub Run_Clear_SearchKey4()
'�����������N���A�@ �J�X�^���ҏW�o�^�p
    Dim Ws As Worksheet
    
    Set Ws = Sheets("�Ǘ��\�ҏW�o�^")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
        .Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = ""
    End With
    Call Re_CostumSt("�Ǘ��\�ҏW�o�^")
'    Call Run_Search_Costumvew("�Ǘ��\�ҏW�o�^")

End Sub
Public Sub Re_CostumSt(Optional ByVal str_Stn As String = "�J�X�^���r���[")
'���ݒ�O�J�X�^���V�[�g�̃f�[�^���Z�b�g
    Dim L_Ws As Worksheet
    
    Set L_Ws = Sheets(str_Stn)
    With L_Ws
        .Unprotect
         .Range("11:20000").Delete
        .Range("B10:GZ20000").ClearContents
        .Range("G:HZ").EntireColumn.AutoFit '16777164
        If str_Stn = "�J�X�^���r���[" Then
            .Range("11:1000").Interior.Color = 13434828
        ElseIf str_Stn = "�Ǘ��\�ҏW�o�^" Then
            .Range("11:1000").Interior.Color = 16777164
        End If
    End With
    Call St_Lock

End Sub

Public Sub Re_Costum_ALLclear()
'���J�X�^���ݒ�̃N���A���Z�b�g
    Dim Ans As Long
    
    Ans = MsgBox("�\�����̃��R�[�h�����" & vbCrLf & _
                            "�擪�T��ȊO�̐ݒ�J�������S�ăN���A����܂�" & vbCrLf & _
                                "��낵���ł���?", vbYesNo + vbInformation, "�t�@�C�i���A���T�[")
    If Ans = vbNo Then
        End
    End If
    ActiveSheet.Unprotect
    ActiveSheet.Range("G5:GS5").ClearContents
    ActiveSheet.Range("G7:GS7").ClearContents
    ActiveSheet.Range("E10:GS80000").ClearContents
    ActiveSheet.Range("B10").Select
    Call Re_CostumSt("�Ǘ��\�ҏW�o�^")
    Call St_Lock

End Sub

Public Sub Clear_SearchKey_G()
'���O���f�[�^���������N���A
    ActiveSheet.Unprotect
    Sheets("�O���f�[�^").Rows(3).ClearContents
    Call Run_Douki_GAIB
    Call St_Lock
    Call Re_Scrl

End Sub
