Attribute VB_Name = "IESample"
Option Explicit

Sub IEOperation()

    Dim ie As IEClass '�N���X�̐錾
    Set ie = New IEClass '�N���X�̎��̉�

    Set ie.objIE = New InternetExplorer
    ie.objIE.Visible = True 'True:IE��\���AFalse:IE���\��
    ie.objIE.Navigate "https://www.yahoo.co.jp/" 'URL�ɓ��ꂽIE���N��
    Call WaitIE(ie.objIE) 'IE�̓ǂݍ��ݑ҂��֐�

    Set ie.htmldoc = ie.objIE.Document '�J����IE�̃h�L�������g���Z�b�g

    'IE���������------------------------


    'IE����I��---------------------------

    ie.objIE.Quit 'IE�����

End Sub

Function WaitIE(objIE As InternetExplorer) 'IE�̓ǂݍ��ݑ҂��֐�

    Do While objIE.Busy = True Or objIE.ReadyState < 4
        DoEvents
    Loop

End Function
