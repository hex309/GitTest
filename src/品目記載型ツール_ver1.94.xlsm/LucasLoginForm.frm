VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LucasLoginForm 
   Caption         =   "LUCAS�F��"
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4395
   OleObjectBlob   =   "LucasLoginForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "LucasLoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Userform_initialize()

    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"

    Dim ret As Long
    Dim formHWnd As Long

    'Get window handle of the userform
    formHWnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)
    'If formHWnd = 0 Then Debug.Print Err.LastDllError

    'Set userform window to 'always on top'
    
    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'If ret = 0 Then Debug.Print Err.LastDllError

 '   Application.WindowState = xlMinimized ' ���̑���͕K�{

End Sub

'�u���O�C���v�{�^���N���b�N��
Private Sub CommandButton1_Click()

 Set PubClsLucasAuth = New LucasAuth
    
    If LucasLoginForm.TextBox1.Value = "" Then Exit Sub
    If LucasLoginForm.TextBox2.Value = "" Then Exit Sub
    
    PubClsLucasAuth.LucasID = LucasLoginForm.TextBox1.Value
    PubClsLucasAuth.LucasPW = LucasLoginForm.TextBox2.Value
    
    Pub�Ј��ԍ� = PubClsLucasAuth.LucasID
    Unload LucasLoginForm
    
End Sub

'�u����v�{�^���N���b�N��
Private Sub CommandButton2_Click()
    Unload LucasLoginForm
    End
End Sub
