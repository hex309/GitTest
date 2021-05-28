Attribute VB_Name = "ImpAces_MOD"
Option Explicit
'��Access�C���|�[�g�n���W���[��

Public Sub ins_Ex_Ac()
'��Excel��AccessDB�O���f�[�^�C���T�[�g 200x2
    Call ins_St_Tbl("TMP_R1$A1:GR", "T_GAIBU1")
    Call ins_St_Tbl("TMP_R2$A1:GR", "T_GAIBU2")

End Sub

Public Sub ins_Ex_Ac_Kanri_CSV()
'��Excel��AccessDB�Ǘ��\�o�b�N�A�b�vCSV�f�[�^�C���T�[�g 202
    Call ins_St_Tbl("TMP_CSV$A1:GT", "T_KANRI", "TMP_CSV", 3, "T_1")

End Sub

Public Function ins_St_Tbl(ByVal str_rng As String, str_Tbl As String, _
                                        Optional str_Stn As String = "�ꊇ�捞", _
                                        Optional RowCnt As Long = 1, _
                                        Optional KeyCol As String = "F_1", Optional Flg As Long = 0)
'��Excel�f�[�^��Access�C���T�[�g
    '���w��V�[�g����w��e�[�u���փf�[�^���C���T�[�g(�����P�V�[�g���Ɣ͈́A�����Q�e�[�u����
    ',�e�[�u���f���[�g�t���O 0=����f���[�g�@1=�f���[�g�����@�f�t�H���g=0)
    Dim i, eRow As Long
    Dim str_Fildn As String
    Dim WHword As String
    
    eRow = Sheets(str_Stn).Cells(Rows.Count, RowCnt).End(xlUp).Row '�ꊇ�捞�V�[�g�ŏI�s�擾
    Call Opn_ExlRs(str_rng & eRow, KeyCol) '�Ǐo�f�[�^�Z�b�g Excel
    Call Opn_AcRs(str_Tbl, KeyCol) '�����f�[�^�Z�b�g Access
    Debug.Print Ac_Rs.State
'    On Error GoTo Era
    If Flg = 0 Then 'Flg=�O�Ńe�[�u������N���A
        Set Ac_Cmd = New ADODB.Command
        str_SQL = ""
        str_SQL = str_SQL & "DELETE FROM " & str_Tbl
        With Ac_Cmd
            .ActiveConnection = Ac_Cn
            .CommandText = str_SQL
            .Execute
        End With
    End If
    Debug.Print Ac_Rs.State

    With Ac_Rs '�f�[�^�㏑�J�n
        Do Until Exl_Rs.EOF
            .AddNew
            For i = 0 To Exl_Rs.Fields.Count - 1
                 str_Fildn = Exl_Rs.Fields(i).Name
                 If str_Fildn = "ImpDate" Then
                    ![ImpDate] = Now
                ElseIf str_Fildn = "RegDate" Then
                    ![RegDate] = Now
                End If
                Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
            Next i
            .Update
            Exl_Rs.MoveNext
        Loop
    End With
    Set Ac_Cmd = New ADODB.Command
    If KeyCol = "F_1" Then
        WHword = "��" '
    ElseIf KeyCol = "T_1" Then
        WHword = "�Ǘ��\ID"
    End If
    str_SQL = ""
    str_SQL = "DELETE FROM " & str_Tbl & " WHERE " & KeyCol & "='" & WHword & "'"  '�f�[�^�^�����ϊ��΍��p���R�[�h�̍폜
    With Ac_Cmd
        .ActiveConnection = Ac_Cn
        .CommandText = str_SQL
        .Execute
    End With
    Call Dis_Ac_Rs
    Call Dis_Exl_Rs
    Exit Function
'�G���[������ ******************************************
Era:
     If Err.Number = -2147467259 Then
        MsgBox "DB�t�@�C���֐ڑ��ł��܂���ł��� " & vbCrLf & _
         "�f�B���N�g���ݒ�Ńp�X���m�F�E�Đݒ肵�Ă�������" & vbCrLf & _
         "OK�������Ɛݒ�y�[�W�ֈړ����܂�", 16
        Call Dis_Ac_Rs
        Call Dis_Exl_Rs
        Call vis_SETDirectSt
        End
    Else
        MsgBox "�G���[" & vbCrLf & _
        Err.Description, 16
        Call Dis_Ac_Rs
        Call Dis_Exl_Rs
        End
    End If

End Function
