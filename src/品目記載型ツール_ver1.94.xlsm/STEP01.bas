Attribute VB_Name = "STEP01"
Option Explicit

Sub STEP01���W���[��()

    Call ���O��������

End Sub

Sub ���O��������()

    Call ����_IE����.ieExistCheck
'    Call ����_�t�H���_�쐬.�e��t�H���_�쐬
    
    If Pub�I�[�g�p�C���b�g�ԍ� = "" Then MsgBox "�I�[�g�p�C���b�g�ԍ�������܂���B�I�����܂��B", vbExclamation: End
    If Pub����V�[�g�� = "" Then MsgBox "����V�[�g��������܂���B�I�����܂��B", vbExclamation: End
    If Pub���σV�[�g�� = "" Then MsgBox "���σV�[�g��������܂���B�I�����܂��B", vbExclamation: End
    If Pub���ϓo�^���� = "" Then MsgBox "������������܂���B�I�����܂��B", vbExclamation: End
    If Pub�H�������d�l�� = "" Then MsgBox "���H�������d�l�����I������Ă��܂���B�I�����܂��B", vbExclamation: End
    
    If Pub�H�������d�l�� = "���Ɩ@" Then
        If Pub�H��FROM = "" Then MsgBox "���Ɩ@�Ώۂ́A���H��FROM���K�v�ł��B�I�����܂��B", vbExclamation: End
        If Pub�H��TO = "" Then MsgBox "���Ɩ@�Ώۂ́A���H��TO���K�v�ł��B�I�����܂��B", vbExclamation: End
        If Pub��C�҃R�[�h = "" Then MsgBox "���Ɩ@�Ώۂ́A����C�҃R�[�h���K�v�ł��B�I�����܂��B", vbExclamation: End
    End If
    
    If Pub�H�������d�l�� <> "�Ȃ�" Then
        If Pub�X�܃R�[�h = "" Then
            MsgBox "���X�܃R�[�h������܂���B�����d�l���쐬�̍ۂɕK�v�ł��B�I�����܂��B", vbExclamation: End
        End If
        Call ����_�H�������d�l���쐬.����_�H�������d�l���쐬���W���[��
    End If
    
End Sub


