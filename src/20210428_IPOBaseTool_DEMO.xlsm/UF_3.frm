VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_3 
   Caption         =   "�O���f�[�^�J����ID�o�^�t�H�[��"
   ClientHeight    =   4515
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11040
   OleObjectBlob   =   "UF_3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
'���t�H�[���N�������X�g�l�Ǎ�&�V�K�o�^ID�Ǎ�
    Call Get_SearchD1
    Me.ListBox1.Clear
    Me.Repaint
    Me.TB_0.SetFocus

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
    
    str_Skey = Me.TB_0.Value
    Call Get_SearchD1(str_Skey)

End Sub

Private Sub CMD_3_Click()
'���폜�{�^���N���b�N
    Dim str_ID0 As String
    Dim Ans As Long
    
    With Me
        str_ID0 = .TB_2.Value
        If str_ID0 = "" Then
            MsgBox "ID�����͂���Ă܂���", 16, "�����̓G���["
            Exit Sub
        End If
    End With
    Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID0 & "'")
    With Ac_Rs
        Set Ac_Cmd = New ADODB.Command
        str_SQL = ""
        str_SQL = str_SQL & "DELETE FROM T_KANRI"
        str_SQL = str_SQL & " WHERE T_1='" & str_ID0 & "'"
        With Ac_Cmd
            .ActiveConnection = Ac_Cn
            .CommandText = str_SQL
            .Execute
        End With
        Ans = MsgBox("����ID�̃f�[�^���Ǘ��\DB����폜����܂�" & vbCrLf & _
                                "�폜���܂���?" & vbCrLf & vbCrLf & _
                                "�͂��@�@�ō폜" & vbCrLf & _
                                "�������@�ŃL�����Z�����܂�", _
                                vbYesNo + vbInformation, "�폜�m�F")
        If Ans = vbYes Then
            MsgBox "�폜���������܂���", vbInformation
        ElseIf Ans = vbNo Then
            MsgBox "�L�����Z������܂���", vbInformation
            Exit Sub
        End If
    End With
    Call Dis_Ac_Rs
    Unload UF_3

End Sub

Private Sub CMD_4_Click()
'���߂�{�^���N���b�N
    Unload UF_3

End Sub

Private Sub ListBox1_Click()
'�����X�g�{�b�N�X�N���b�N�C�x���g
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 0)
    End With
    
End Sub

Public Function Get_SearchD1(Optional ByVal str_Skey As Variant = "")
'���������e�Ń��R�[�h�Z�b�g�����˃��X�g�{�b�N�X���f�@�Ǘ��\ID�p
    Dim str_SQL As String
 '�Ǐo�f�[�^�Z�b�g
     If str_Skey <> "" Then
        str_SQL = str_SQL & " AND T_1 LIKE'%" & str_Skey & "%'"
    End If
    Call Opn_AcRs("T_KANRI", "T_1", str_SQL)
 '���X�g�{�b�N�X�ɒǉ�
    With Me.ListBox1
        .Clear
        Do Until Ac_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(Ac_Rs!T_1), "", Ac_Rs!T_1)
            Ac_Rs.MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

End Function
