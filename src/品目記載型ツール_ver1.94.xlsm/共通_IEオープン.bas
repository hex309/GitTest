Attribute VB_Name = "����_IE�I�[�v��"
Option Explicit

Public Sub �V�KIE�J��(ByVal URL As String)
    
    Call ieOpenNew(URL)

End Sub

Public Sub ����IE�J��(ByVal URL As String)

    Call ieOpenExist(URL)

End Sub

Public Sub ieOpenNew(urlName As String, _
           Optional viewFlg As Boolean = True, _
           Optional ieTop As Integer = 30, _
           Optional ieLeft As Integer = 0, _
           Optional ieWidth As Integer = 900, _
           Optional ieHeight As Integer = 900)
    
    'IE(InternetExplorer)�̃I�u�W�F�N�g���쐬����
    'InternetExplorer.Application ��.navigate���\�b�h���s��ɁA�C���X�^���X�I�u�W�F�N�g���p�������B
    '�Z�L�����e�B�]�[�����܂����ꍇ�̃Z�b�V�������ێ��ł��Ȃ����߁B
    'Set�����LowL�A�ی샂�[�h�I�t�� .navigate���\�b�h�́AMidiumL�̂��߁A�C���X�^���X�������p���Ȃ�
    'Set oPubIE1 = CreateObject("InternetExplorer.Application")�@'LowL

    '��L�̑��
    Set oPubIE1 = New InternetExplorerMedium
    
    With oPubIE1
        
        'IE(InternetExplorer)��\���E��\��
        .Visible = viewFlg
        
        .Top = ieTop  'Y�ʒu
        .Left = ieLeft  'X�ʒu
        .Width = ieWidth  '��
        .Height = ieHeight  '����
        
        '�w�肵��URL�̃y�[�W��\������
        .navigate urlName
    
    End With

    'IE(InternetExplorer)�����S�\�������܂őҋ@
    Call ieWaitCheck
    
End Sub

Sub ieOpenExist(ByVal urlName As String)

    Const navOpenInNewTab = &H800

    Dim objShell As Object, objWin As Object
    Dim nFLG As Boolean: nFLG = False

    Set objShell = CreateObject("Shell.Application")

    For Each objWin In objShell.Windows

        '���݂��Ȃ��̂ɁAInternet Explorer ���\�������Ƃ�������B����SUB�͐������Ȃ�
        If objWin.Name = "Internet Explorer" Then
            Set oPubIE1 = objWin
            nFLG = True
            Exit For
        End If
    Next
    
    If nFLG = True Then
        oPubIE1.Navigate2 urlName, navOpenInNewTab
    Else
        ieOpenNew (urlName)
    End If

End Sub


