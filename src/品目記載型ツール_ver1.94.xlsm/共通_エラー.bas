Attribute VB_Name = "����_�G���["
Option Explicit

'-------------------------
'�G���[����
'-------------------------
Public Sub IE�s���S����G���[(ByVal errWin As String)
    
    Dim ���b�Z�[�W As String
    
    ���b�Z�[�W = "����ԍ�: " & Pub����ԍ� & vbLf & vbLf & "ErrTitleWin: " & errWin

    MessageBox 0, ���b�Z�[�W, "�I�[�g�p�C���b�g��~", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION

    End

End Sub
