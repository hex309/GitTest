VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �{���{�b�N�X 
   Caption         =   "���ϑO�����"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "�{���{�b�N�X.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�{���{�b�N�X"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Mod�V�[�g�� As String

Private Sub Userform_initialize()

    Mod�V�[�g�� = ThisWorkbook.ActiveSheet.Name

    TextBox1.Value = Replace(ThisWorkbook.Sheets(Mod�V�[�g��).Range(Pub�{���A�h���X).Value, vbCr, "")
    
End Sub

'-----------------------------------------
' �u����v�{�^��
'-----------------------------------------
Private Sub CommandButton1_Click()
        
    ThisWorkbook.Sheets(Mod�V�[�g��).Range(Pub�{���A�h���X).Value = TextBox1.Value
    
    Unload Me

End Sub
