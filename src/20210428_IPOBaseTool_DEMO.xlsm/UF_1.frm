VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_1 
   Caption         =   "�O���f�[�^�J����ID�o�^�t�H�[��"
   ClientHeight    =   7605
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8565
   OleObjectBlob   =   "UF_1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'���t�H�[���N�������X�g�l�Ǎ�
    Call Get_SearchD

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'��X�ŕ����Ȃ�����
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
End Sub

Private Sub CMD_1_Click()
'�������{�^���N���b�N
    Dim str_Skey As Variant
    
    str_Skey = Me.TB_1.Value
    Call Get_SearchD(str_Skey)

End Sub

Private Sub CMD_2_Click()
'���o�^�{�^���N���b�N
    Dim eRow As Long
    Dim Ws As Worksheet
    Dim str_Ans As String
    
    Set Ws = Sheets("�J�����ݒ�")
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
    End With
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "�Ǘ��\�J����ID��ݒ肵�Ă���s���Ă�������", 16, "�Ǘ��\�J���������̓G���["
            Exit Sub
        End If
        .Cells(eRow, 7).Value = str_Ans
    End With
    Unload UF_1


End Sub

Private Sub CMD_3_Click()
'���߂�{�^���N���b�N
    Unload UF_1

End Sub

Private Sub ListBox1_Click()
'�����X�g�{�b�N�X�N���b�N�C�x���g
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 1)
    End With
    
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'�����X�g�{�b�N�X�_�u���N���b�N�C�x���g
    Dim eRow As Long
    Dim Ws As Worksheet
    Dim str_Ans As String
    
    Set Ws = Sheets("�J�����ݒ�")
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
    End With
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "�Ǘ��\�J����ID��ݒ肵�Ă���s���Ă�������", 16, "�Ǘ��\�J���������̓G���["
            Exit Sub
        End If
        .Cells(eRow, 7).Value = str_Ans
    End With
    Unload UF_1

End Sub

Public Function Get_SearchD(Optional ByVal str_Skey As Variant = "")
'���������e�Ń��R�[�h�Z�b�g�����˃��X�g�{�b�N�X���f
    Const adOpenKeyset = 1, adLockReadOnly = 1
    Dim str_RCn  As String
    Dim R_Cn As ADODB.Connection
    Dim R_Rs As ADODB.Recordset
    Dim str_SQL As String
    Dim eRow As Long
    Dim R_Ws As Worksheet

    Set R_Ws = Sheets("T_GAIBColList")
    eRow = R_Ws.Cells(Rows.Count, 5).End(xlUp).Row
 '�Ǐo�f�[�^�Z�b�g *******************************************************************
    Set R_Cn = New ADODB.Connection
    Set R_Rs = New ADODB.Recordset
    If R_Cn.State = 1 Then End
    R_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    R_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=NO;IMEX=1"
    R_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT * "
    str_SQL = str_SQL & " FROM [T_GAIBColList$A3:B500] "
    If str_Skey <> "" Then
        str_SQL = str_SQL & " WHERE F2 LIKE'%" & str_Skey & "%'"
    End If
    
    R_Rs.Open str_SQL, R_Cn, adOpenKeyset, adLockReadOnly
 '�Ǐo�f�[�^�Z�b�g�����܂� **************************************************************
 '���X�g�{�b�N�X�ɒǉ�
    With Me.ListBox1
        .Clear
        Do Until R_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(R_Rs!F2), "", R_Rs!F2)
            .List(.ListCount - 1, 1) = R_Rs!F1
            R_Rs.MoveNext
        Loop
    End With
'���㏈��
    R_Rs.Close '���R�[�h�Z�b�g�̃N���[�Y
    Set R_Rs = Nothing
    R_Cn.Close '�R�l�N�V�����̃N���[�Y
    Set R_Cn = Nothing  '�I�u�W�F�N�g�̔j��

End Function



