VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_2 
   Caption         =   "�O���f�[�^�J����ID�o�^�t�H�[��"
   ClientHeight    =   7400
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13680
   OleObjectBlob   =   "UF_2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'���t�H�[���N�������X�g�l�Ǎ�&�V�K�o�^ID�Ǎ�
    
    Call Get_SearchD1
    Call Get_SearchD2
    Me.TB_2.Value = Sheets("�Ǘ��\�V�K�o�^").Range("D6").Value
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

Private Sub CMD_2_Click()
'�������{�^���N���b�N
    Dim str_Skey As Variant
    
    str_Skey = Me.TB_5.Value
    Call Get_SearchD2(str_Skey)

End Sub

Private Sub CMD_3_Click()
'���o�^�{�^���N���b�N
    Dim str_ID0 As String
    Dim str_ID1 As String
    Dim str_ID2 As String
    Dim Ans As Long
    
    With Me
        str_ID0 = .TB_2.Value
        str_ID1 = .TB_3.Value
        str_ID2 = .TB_4.Value
        If .TB_2.Value = "" Or .TB_3.Value = "" Or .TB_4.Value = "" Then
            MsgBox "�R�tID����S�Đݒ肵�Ă���s���Ă�������", 16, "�����̓G���["
            Exit Sub
        End If
    End With
    Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID0 & "'")
    With Ac_Rs
        If IIf(IsNull(!T_2), "", !T_2) <> "" Then
            Ans = MsgBox("���̊Ǘ��\ID�ɂ͊��ɕR�tID�o�^����Ă��܂���" & vbCrLf & _
                                    "�\������ID�ŏ㏑���o�^���܂���?" & vbCrLf & vbCrLf & _
                                    "�͂��@�@�ŏ㏑��" & vbCrLf & _
                                    "�������@�ŃL�����Z�����܂�", _
                                    vbYesNo + vbInformation, "���ɕR�tID���o�^����Ă��܂�")
            If Ans = vbYes Then
                !T_2 = str_ID1
                !T_3 = str_ID2
                .Update
                MsgBox "�R�t�o�^���������܂���", vbInformation
                Ans = MsgBox("�o�^���ʂ��m�F���܂����H" & vbCrLf & _
                "�͂� �@ �ŊǗ��\�ҏW��ʂ�" & vbCrLf & _
                "�������@�Ńz�[����ʂɈړ����܂�", vbYesNo + vbInformation, "���ʂ̊m�F")
                If Ans = vbYes Then
                    Call vis_KANRISt
                    With Sheets("�Ǘ��\�ҏW�o�^")
                        .Range("D4").Value = str_ID0
                        .Range("E4").Value = str_ID1
                        .Range("F4").Value = str_ID2
                    End With
                    Call Run_Search_Costumvew("�Ǘ��\�ҏW�o�^")
                    Call Re_Scrl
                ElseIf Ans = vbNo Then
                    Call vis_UISt
                End If
            ElseIf Ans = vbNo Then
                MsgBox "�L�����Z������܂���", vbInformation
                Exit Sub
            End If
        Else
            !T_2 = str_ID1
            !T_3 = str_ID2
            .Update
            MsgBox "�R�t�o�^���������܂���", vbInformation
            Ans = 0
            Ans = MsgBox("�o�^���ʂ��m�F���܂����H" & vbCrLf & _
            "�͂� �@ �ŊǗ��\�ҏW��ʂ�" & vbCrLf & _
            "�������@�Ńz�[����ʂɈړ����܂�", vbYesNo + vbInformation, "���ʂ̊m�F")
            If Ans = vbYes Then
                Call vis_KANRISt
                With Sheets("�Ǘ��\�ҏW�o�^")
                    .Unprotect
                    .Range("D4").Value = str_ID0
                    .Range("E4").Value = str_ID1
                    .Range("F4").Value = str_ID2
                End With
                Call Run_Search_Costumvew("�Ǘ��\�ҏW�o�^")
                Call Re_Scrl
            ElseIf Ans = vbNo Then
                Call vis_UISt
            End If
    End If

    End With
    Call Dis_Ac_Rs
    Unload UF_2

End Sub

Private Sub CMD_4_Click()
'���߂�{�^���N���b�N
    Unload UF_2

End Sub

Private Sub ListBox1_Click()
'�����X�g�{�b�N�X�N���b�N�C�x���g
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 0)
    End With
    
End Sub

Private Sub ListBox2_Click()
'�����X�g�{�b�N�X�N���b�N�C�x���g
    With Me.ListBox2
        Me.TB_3.Value = .List(.ListIndex, 0)
        Me.TB_4.Value = .List(.ListIndex, 1)
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

Public Function Get_SearchD2(Optional ByVal str_Skey As Variant = "")
'���������e�Ń��R�[�h�Z�b�g�����˃��X�g�{�b�N�X���f�@�O���f�[�^ID�p
    Dim str_SQL As String
 '�Ǐo�f�[�^�Z�b�g
     If str_Skey <> "" Then
        str_SQL = str_SQL & " AND F_1 LIKE'%" & str_Skey & "%'"
    End If
    Call Opn_AcRs("T_GAIBU1", "F_1", str_SQL)
 '���X�g�{�b�N�X�ɒǉ�
    With Me.ListBox2
        .Clear
        Do Until Ac_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(Ac_Rs!F_1), "", Ac_Rs!F_1)
            .List(.ListCount - 1, 1) = IIf(IsNull(Ac_Rs!F_2), "", Ac_Rs!F_2)
            Ac_Rs.MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

End Function
