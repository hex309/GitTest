VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AlertBox 
   Caption         =   "�i���m�F"
   ClientHeight    =   1980
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   4305
   OleObjectBlob   =   "AlertBox.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "AlertBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cancelBtn_Click()
    cancelFlg = True
End Sub

'@1
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'�o�c���΍�

    If Not CloseMode = vbFormCode Then
        Cancel = True
    End If
End Sub

'���[�U�[�t�H�[����������ƁA�V�[�g�̕ҏW���ł��Ȃ��Ȃ�o�O�i2013�ŗL�H�j���������
'https://support.microsoft.com/ja-jp/help/2851316
'���[�U�[�t�H�[�����z�u�A�Ĕz�u�����ƌ��m

Private Sub UserForm_Layout()
    '�ÓI�ϐ��Ő錾���A2�x�ڈȍ~�̓C�x���g���m���Ă������͍s��Ȃ��悤�ɂ���
    Static fSetModal As Boolean
    If fSetModal = False Then
        fSetModal = True
        '�t�H�[�����\����
        Me.Hide
        '�t�H�[�������[�_���ŕ\��
        Me.Show vbModeless
        
        '�t�H�[����`�悳����]�T�����
        DoEvents

    End If
End Sub


