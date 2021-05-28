Attribute VB_Name = "Upd_MOD"
Option Explicit
'��SQL�A�b�v�f�[�g�n���W���[��

Public Sub upd_NewID()
'���V�KID�o�^�X�V
    Dim str_Cval As String
    
    str_Cval = Sheets("�Ǘ��\�V�K�o�^").Range("D6").Value
    If CHK_Duplicate_ID("T_KANRI", str_Cval, "T_1") = True Then
        MsgBox "����ID�͊��Ɏg���Ă��܂�", 16, "�d���G���[!"
        End
    End If
    Call Opn_AcRs("T_KANRI", "T_1")
    With Ac_Rs
        .AddNew
        !T_1 = str_Cval
        !RegDate = Now
        .Update
    End With
    Call Dis_Ac_Rs
  
End Sub

Public Sub Run_updKANRI()
'���Ǘ��\�o�^�X�V
    Call upd_KANRI
    
End Sub

Public Sub Run_updCostum_KANRI()

    Call upd_KANRI("�Ǘ��\�ҏW�o�^")
    
End Sub

Public Sub upd_KANRI(Optional ByVal str_Stn As String = "�Ǘ��\�ҏW�o�^", Optional Flg As Long = 0)
'���f�[�^�㏑���X�V����
    '�H���Ǘ��\�̍X�V�L��='�L'�̂�(T_KANRI)�㏑�X�V
    '�O���f�[�^�e�[�u��(T_GAIBU1,2)�ւ͍X�V�L��='�L'�̂�2�L�[���j�[�N(F_1&F_2)�ŏ㏑
    Dim str_Fildn, str_Ans As String
    Dim i As Long
    Dim str_RngAd As String
    
    Call CHK_RegChange '�X�V�L������
    str_Ans = Get_ChangeData '�X�V�f�[�^ID�擾
    str_RngAd = Sheets(str_Stn).Range("B7").End(xlToRight).Address
    str_RngAd = Replace(str_RngAd, "7", "50")
    str_RngAd = Replace(str_RngAd, "$", "")
    Call Opn_ExlRs(str_Stn & "$B7:" & str_RngAd, "T_1", " AND RegFlg='�L'")
    Call Opn_AcRs("T_KANRI", "T_1")
    With Ac_Rs '�Ǘ��\�㏑�X�V�J�n
        Do Until .EOF
            If !T_1 = Exl_Rs!T_1 Then
                For i = 1 To Exl_Rs.Fields.Count - 1
                    str_Fildn = Exl_Rs.Fields(i).Name
                    Ac_Rs![RegFlg] = "�X�V�L"
                    Ac_Rs![RegDate] = Now
                    Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
                    .Update
                Next i
            End If
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Call Dis_Ac_Rs
    Call upd_GAIB '�O���f�[�^�e�[�u���̏㏑�X�V
    Call Run_Search_KANRI '�f�[�^�ē�
    MsgBox "�f�[�^���X�V����܂��� !!" & vbCrLf & _
                  str_Ans, vbInformation
    Exit Sub
Era:
     
End Sub

Public Function upd_GAIB()
'���O���f�[�^�̏㏑���X�V
    Dim str_Fildn As String
    Dim str_SQLFild As String
    Dim i As Long
    
    Call Opn_ExlRs("�Ǘ��\�ҏW�o�^$B6:FZ8000", "F_1", " AND RegFlg='�L'")
    Call Opn_AcRs("T_GAIBU1", "F_1")
    With Ac_Rs
        Do Until .EOF
            If !F_1 = Exl_Rs!F_1 Then
                If !F_2 = Exl_Rs!F_2 Then
                    For i = 0 To Exl_Rs.Fields.Count - 1
                        On Error GoTo Skip0
                        str_Fildn = Exl_Rs.Fields(i).Name
Skip0:
                        Ac_Rs![RegFlg] = "�X�V�L"
                        If str_Fildn = "ID" Then GoTo Skip1  '�V�[�g�ɂȂ��J�����̓X�L�b�v
                        If str_Fildn = "ImpDate" Then GoTo Skip1 '�V�[�g�ɂȂ��J�����̓X�L�b�v
                        If str_Fildn = "RegFlg" Then GoTo Skip1 '�V�[�g�ɂȂ��J�����̓X�L�b�v
                        If InStr(str_Fildn, "_") <= 0 Then GoTo Skip1
                        Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
Skip1:
                    Next i
                    .Update
                End If
            End If
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Call Dis_Ac_Rs

End Function

Public Sub upd_KANRI_RegFlgRe()
'���Ǘ��\�e�[�u���̍X�V�L����S�ă��Z�b�g
    '�C���|�[�g���Ɏg�p
    Dim sr_SQL As String
    
    Call Opn_AcRs("T_KANRI", "T_1")
    str_SQL = ""
    str_SQL = "UPDATE T_KANRI SET RegFlg=''"
    Set Ac_Cmd = New ADODB.Command
    With Ac_Cmd
        .ActiveConnection = Ac_Cn
        .CommandText = str_SQL
        .Execute
    End With
    Call Dis_Ac_Rs
    
End Sub
