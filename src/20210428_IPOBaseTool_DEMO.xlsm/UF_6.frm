VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_6 
   Caption         =   "�\�����������ڂ�I��ł�������"
   ClientHeight    =   4750
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12270
   OleObjectBlob   =   "UF_6.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
'���폜�{�^���N���b�N
    Dim Ans As Long
    Dim aRow As Long
    Dim str_Ans As String
    Dim R_Ws As Worksheet
    
    If Me.TB_2.Value = "" Then Exit Sub
    Set R_Ws = Sheets("�J�X�^���ҏW�o�^���C�ɓ���")
    Ans = MsgBox("�I�𒆂̓o�^�����폜���܂�" & vbCrLf & _
                        "��낵���ł����H", vbYesNo + vbInformation, "�폜���܂���")
    If Ans = vbNo Then End
    str_Ans = Me.TB_2.Value
    With R_Ws
        .Unprotect
        aRow = Application.WorksheetFunction.Match(str_Ans, .Range("A1:A1000"), 0)
        .Range(aRow & ":" & aRow).Delete
    End With
    Me.TB_2.Value = ""
    MsgBox "�폜����܂���", vbInformation
    Call Get_SearchD
    Call St_Lock

End Sub

Private Sub UserForm_Activate()

    Call Get_SearchD

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
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
'���ďo�{�^���N���b�N
    Dim str_Ans As String
    Dim aRow As Long
    Dim R_Ws As Worksheet
    
    Set R_Ws = Sheets("�J�X�^���ҏW�o�^���C�ɓ���")
    str_Ans = Me.TB_2.Value
    If str_Ans = "" Then
        MsgBox "�ďo�o�^�����I������Ă��܂���", 16
        Exit Sub
    End If
    ActiveSheet.Unprotect
    With R_Ws
        aRow = Application.WorksheetFunction.Match(str_Ans, .Range("A1:A1000"), 0)
        .Range(.Cells(aRow, 2), .Cells(aRow, 200)).Copy
    End With
    ActiveSheet.Range("G7").PasteSpecial Paste:=xlValues
    ActiveSheet.Range("G:HZ").EntireColumn.AutoFit
    Unload UF_6
    Call St_Lock

End Sub

Private Sub CMD_3_Click()
'������{�^���N���b�N
    Unload UF_6

End Sub

Private Sub ListBox1_Click()
'�����X�g�{�b�N�X�N���b�N�C�x���g
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 0)
    End With
    
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'�����X�g�{�b�N�X�_�u���N���b�N�C�x���g
    Dim str_Ans As String
    Dim aRow As Long
    Dim R_Ws As Worksheet
    
    Set R_Ws = Sheets("�J�X�^���ҏW�o�^���C�ɓ���")
    With Me.ListBox1
       str_Ans = .List(.ListIndex, 0)
    End With
    ActiveSheet.Unprotect
    With R_Ws
        aRow = Application.WorksheetFunction.Match(str_Ans, .Range("A1:A1000"), 0)
        .Range(.Cells(aRow, 2), .Cells(aRow, 200)).Copy
    End With
    ActiveSheet.Range("G7").PasteSpecial Paste:=xlValues
    ActiveSheet.Range("G:HZ").EntireColumn.AutoFit
    Unload UF_6
    Call St_Lock

End Sub

Public Function Get_SearchD(Optional ByVal str_Skey As Variant = "")
'���������e�Ń��R�[�h�Z�b�g�����˃��X�g�{�b�N�X���f
    Const adOpenKeyset = 1, adLockReadOnly = 1
    Dim str_RCn  As String
    Dim R_Cn As ADODB.Connection
    Dim R_Rs As ADODB.Recordset
    Dim str_SQL As String
 '�Ǐo�f�[�^�Z�b�g *******************************************************************
    Set R_Cn = New ADODB.Connection
    Set R_Rs = New ADODB.Recordset
    If R_Cn.State = 1 Then End
    R_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    R_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=YES;IMEX=1"
    R_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT * "
    str_SQL = str_SQL & " FROM [�J�X�^���ҏW�o�^���C�ɓ���$A1:B500] "
    If str_Skey <> "" Then
        str_SQL = str_SQL & " WHERE �o�^�� LIKE'%" & str_Skey & "%'"
    End If
    
    R_Rs.Open str_SQL, R_Cn, adOpenKeyset, adLockReadOnly

 '�Ǐo�f�[�^�Z�b�g�����܂� **************************************************************
 '���X�g�{�b�N�X�ɒǉ�
    With Me.ListBox1
        .Clear
        Do Until R_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(R_Rs!�o�^��), "", R_Rs!�o�^��)
            R_Rs.MoveNext
        Loop
    End With
'���㏈��
    R_Rs.Close '���R�[�h�Z�b�g�̃N���[�Y
    Set R_Rs = Nothing
    R_Cn.Close '�R�l�N�V�����̃N���[�Y
    Set R_Cn = Nothing  '�I�u�W�F�N�g�̔j��

End Function
