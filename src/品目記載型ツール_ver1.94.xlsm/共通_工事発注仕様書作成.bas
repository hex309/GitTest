Attribute VB_Name = "����_�H�������d�l���쐬"
Option Explicit

Dim Mod�A�h�I���t���p�X As String
Dim Mod�A�h�I�����s���W���[�� As String
Dim Mod�c�[���t���p�X As String
    
Sub ����_�H�������d�l���쐬���W���[��()
    
    Call �A�h�I�������ݒ�
    Call �A�h�I�����s

End Sub

Sub �A�h�I�������ݒ�()

    Mod�A�h�I���t���p�X = ThisWorkbook.Path & "\" & "�yadd-on�zSSIS�����o�^.xlam"
    Mod�A�h�I�����s���W���[�� = "���C��_�H�������d�l���쐬.���C��_�H�������d�l���쐬���W���[��"
    
    Mod�c�[���t���p�X = ThisWorkbook.FullName
    
End Sub

Sub �A�h�I�����s()
    
    Dim �A�h�I���t���p�X As String
    Dim �A�h�I�����s���W���[�� As String
    Dim �A�h�I���r����~�t���O As Boolean
    Dim �c�[���t���p�X As String
   
    '-----------------------------------------------
    ' �A�h�I���t�@�C���́A����K�w�ɑ��݂��邱�ƁB
    '-----------------------------------------------
    �A�h�I���t���p�X = Mod�A�h�I���t���p�X
    �A�h�I�����s���W���[�� = Mod�A�h�I�����s���W���[��
    �c�[���t���p�X = Mod�c�[���t���p�X
    
    '-----------------------------------------------
    ' �A�h�I�����݊m�F
    '-----------------------------------------------
    If Dir(�A�h�I���t���p�X) = "" Then
        MsgBox �A�h�I���t���p�X & "�����݂��Ȃ����ߒ��~���܂�", vbExclamation
        End
    End If
    
    '-----------------------------------------------
    ' �A�h�I�����s
    '-----------------------------------------------
    Dim strJoin As String
    strJoin = "'" & �A�h�I���t���p�X & "'!" & �A�h�I�����s���W���[��
    
    Application.Run strJoin, �c�[���t���p�X, Pub���σV�[�g��, Pub�H�������d�l��, Pub�X�܃R�[�h, Pub���ϓo�^����

End Sub


