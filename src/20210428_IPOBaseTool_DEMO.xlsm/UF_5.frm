VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_5 
   Caption         =   "���C�ɓ���ۑ�"
   ClientHeight    =   4250
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5775
   OleObjectBlob   =   "UF_5.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
End Sub

Private Sub CMD_2_Click()
'���o�^�{�^���N���b�N
    Dim eRow As Long
    Dim str_Ans As String

    str_Ans = Me.TB_2.Value
    If str_Ans = "" Then
        MsgBox "�o�^�������͂���Ă��܂���", 16
        Exit Sub
    End If
    With Sheets("�J�X�^���ҏW�o�^���C�ɓ���")
        eRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
        Call Opn_ExlRs("�J�X�^���ҏW�o�^���C�ɓ���$A1:B1000", "�o�^��")
        With Exl_Rs
            Do Until Exl_Rs.EOF
               If !�o�^�� = str_Ans Then GoTo Skip
                .MoveNext
            Loop
        End With
        Call Dis_Exl_Rs
        .Unprotect
        .Cells(eRow, 1).Value = str_Ans
        Sheets("�Ǘ��\�ҏW�o�^").Range("G7:GU7").Copy
        .Cells(eRow, 2).PasteSpecial Paste:=xlValues
    End With
    ThisWorkbook.Save
    MsgBox "�o�^�����I", vbInformation
    Unload UF_5
    Call St_Lock
    
    Exit Sub
Skip:
    MsgBox "���̖��O�͊��Ɏg���Ă��܂�", 16
    Call Dis_Exl_Rs
    Exit Sub

End Sub

Private Sub CMD_3_Click()
'������{�^���N���b�N
    Unload UF_5

End Sub
