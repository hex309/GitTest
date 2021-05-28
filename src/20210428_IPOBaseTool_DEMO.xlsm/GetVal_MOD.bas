Attribute VB_Name = "GetVal_MOD"
Option Explicit
'���l�擾�n���W���[��

Public Function Get_Maxval() As String
'��ID�ő�l�擾
    Dim str_Ans As Variant
    Dim cnt As Long
    
    Sheets("CHK_MID").Range("A2:A80000").ClearContents
    Call Opn_AcRs("T_KANRI", "T_1", , "T_1")
    With Ac_Rs
        cnt = 1
        Do Until .EOF
            str_Ans = ""
            str_Ans = !T_1
            str_Ans = Mid(str_Ans, 4, Len(str_Ans))
            cnt = cnt + 1
            Sheets("CHK_MID").Cells(cnt, 1).Value = str_Ans
            .MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

    str_Ans = ""
    str_Ans = Sheets("CHK_MID").Range("B1").Value
    Debug.Print str_Ans
    Get_Maxval = "XXX" & str_Ans + 1

End Function

Public Function Get_ChangeData()
'���X�V�������������b�Z�[�W�f�[�^�ڍו��쐬
    '�Ǘ��\�㏑���X�V�����O�Ɏg�p
    Dim str_SKey1, str_SKey2, str_SKey3, str_Ans As String
    
    Call Opn_ExlRs("�Ǘ��\�ҏW�o�^$B7:FZ8000", "T_1", " AND RegFlg='�L'")
    If Exl_Rs.EOF = True Then ''�L'�f�[�^���Ȃ������ꍇ
        MsgBox "�ύX���ꂽ�f�[�^�͂���܂���", vbInformation
        Call Dis_Exl_Rs
        End
    End If
    Do Until Exl_Rs.EOF '�Ǐo���f�[�^����X�V���b�Z�[�W�쐬
        str_SKey1 = str_SKey1 & Exl_Rs!T_1 & vbCrLf
        str_SKey2 = str_SKey2 & Exl_Rs!T_2 & "," & Exl_Rs!T_3 & vbCrLf
        Exl_Rs.MoveNext
    Loop
    str_Ans = str_Ans & "�X�V���ꂽ���R�[�h��" & vbCrLf & "�y�Ǘ��\�L�[�z" & vbCrLf
    str_Ans = str_Ans & str_SKey1 & "�y�O���f�[�^�Q�L�[�z" & vbCrLf
    str_Ans = str_Ans & str_SKey2 & "�ł���"
    Call Dis_Exl_Rs
    Get_ChangeData = str_Ans

End Function

Public Function Get_FilFol(ByVal Flg As Long, Optional F_type As String = "") As Variant
'���_�C�A���O����t�@�C��/�t�H���_�p�X�擾�t�@���N�V����
    '(����1:�t�@�C��/�t�H���_�I���t���O 1=�t�@�C���@2=�t�H���_,����2:���b�Z�[�W�ƃt�@�C���g���q�j
    'F_type=�T���v��:"�C���|�[�g�t�@�C����I�����Ă������� (*.xlsb;*.xlsx;*.accdb), *.xlsb;*.xlsx;*.accdb"
    Dim Sfile As String
    Dim i As Integer
    Dim s As String
   
    If Flg = 1 Then
        Sfile = Application.GetOpenFilename(F_type)
        If Sfile = "False" Then
            Get_FilFol = ""
            Exit Function
        End If
            Get_FilFol = Sfile
    ElseIf Flg = 2 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            If Sfile = "False" Then
                Get_FilFol = ""
            Else
                Get_FilFol = .SelectedItems(1)
                
            End If
        End With
    End If

End Function
