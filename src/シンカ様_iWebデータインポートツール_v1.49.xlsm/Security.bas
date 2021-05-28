Attribute VB_Name = "Security"
Option Explicit
Option Private Module

'�u���C���v�V�[�g�̃{�^�������A�u�ی�����v����u3���ԕی�����v�֕ύX
'�E���̃��b�Z�[�W���u���������̓f�[�^�C���|�[�g�͎��s�ł��܂���v�ɂ���
Public Function ReProtect(ByVal sheetName As String, ByVal endTime As Date)
    
    On Error Resume Next
    '����(endTime)�̎����ɂȂ�����A���v���V�[�W���̏��������s
    Application.OnTime endTime, "'ReProtect""" & Worksheets(sheetName).name & """,""" & endTime & """ '", , False
    On Error GoTo 0  '�G���[�𖳌��ɂ���
    
    '�{�^�������u3���ԕی�����v�֕ύX
    '�E���̃��b�Z�[�W���u���������̓f�[�^�C���|�[�g�͎��s�ł��܂���v�ɂ���
    Worksheets(sheetName).protectSheet
    
    Application.EnableEvents = True

    MsgBox Worksheets(sheetName).name & "�V�[�g�̕ی���ĊJ���܂����B", vbInformation

End Function

'�V�[�g�ی��������Password�����������͂��ꂽ�ꍇ�A�{�^�����Ƀt�H���g�F�ԂŁu�ی�����v�I�����Ԃ𖾎�
'�{�^�������u3���ԕی�����v����u�ی�ĊJ�v�֕ύX
Public Sub unprotectFewMinutes(ByVal sheetName As String, ByVal endTime As Date, Optional ByVal dspMsgAdd As String)
    Dim sh As Shape

    On Error GoTo Error  'Password��������ꍇ
    Worksheets(sheetName).Unprotect  '�V�[�g�ی���������邽�߁APassword���͉�ʂ�\��
    
    On Error GoTo 0

    'Password���͉�ʂŁA�u�~�v�܂��́u�L�����Z���v���������ꂽ�ꍇ�́A�����I��
    If Worksheets(sheetName).ProtectContents Then
        Exit Sub
    End If

    '����(endTime)�̎��ԂɂȂ�����A�ی�ĊJ�������s���AFunction�v���V�[�W���uReProtect�v(����)�����s
    Application.OnTime endTime, "'ReProtect""" & Worksheets(sheetName).name & """,""" & endTime & """ '", , True
    Application.EnableEvents = False
    
    If dspMsgAdd <> vbNullString Then  '����(dspMsgAdd)�ɒl������ꍇ�́A�l�Ƀt�H���g�F�Ԃŏ���
        With Worksheets(sheetName).Range(dspMsgAdd)
            .Value = "* " & Format(endTime, "hh��mm��ss�b") & " �ɕی���ĊJ���܂��B"
            .Font.Color = vbRed
        End With
    End If

    For Each sh In Worksheets(sheetName).Shapes  '�ی�{�^���̏ꍇ�A�{�^����������
        If sh.name = "ProtectBtn" Then
             sh.TextEffect.Text = "�ی�ĊJ"
        End If
    Next

    MsgBox "�ی���������܂����B" & vbCrLf & Format(endTime, "hh��mm��ss�b") & " �ɕی���ĊJ���܂�", vbExclamation

Exit Sub
Error:
    MsgBox "�p�X���[�h���Ⴂ�܂��B"

End Sub

Public Function sendCaptAlert(ByVal msgBody As String) As Boolean
    Dim userAcc As String
    Dim sndAdd As String
    Dim msgSubject As String
    
    If ScenarioSh.getUserName = vbNullString Then Exit Function
    
    With MailSettingSh
        userAcc = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = sndAdd & "; " & .getSendAccount(, "�F��/�ǉ�")
        msgSubject = .getCaptSubject
    End With
    
     sendCaptAlert = sendMail(userAcc, sndAdd, msgSubject, msgBody)
End Function

Public Function sendSemAlert(ByVal msgBody As String) As Boolean
    Dim userAcc As String
    Dim sndAdd As String
    Dim msgSubject As String
    
    If ScenarioSh.getUserName = vbNullString Then Exit Function
    
    With MailSettingSh
        userAcc = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = sndAdd & "; " & .getSendAccount(, "�F��/�ǉ�")
        msgSubject = .getSemSubject
    End With
    
     sendSemAlert = sendMail(userAcc, sndAdd, msgSubject, msgBody)
End Function


Public Function sendFinAlert(ByVal msgBody As String, Optional tgtCorp As String = vbNullString) As Boolean
    Dim userAcc As String
    Dim sndAdd As String
    Dim msgSubject As String
    
    With MailSettingSh
        userAcc = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = .getSendAccount(, tgtCorp)
        msgSubject = IIf(tgtCorp = vbNullString, vbNullString, "�y" & tgtCorp & "�z") & .getFinSubject
    End With
    
     sendFinAlert = sendMail(userAcc, sndAdd, msgSubject, msgBody)
End Function


Private Function sendMail(ByVal sendAccount As String, _
                         ByVal toAdd As String, _
                         ByVal subject As String, _
                         ByVal msgBody As String, _
                         Optional ByVal ccAdd As String = vbNullString) As Boolean
                         
    Dim objOl As Object 'Outlook.Application
    Dim objMl As Object 'Outlook.MailItem
    Dim Account As Object ' Outlook.Account
    Dim tgtAcc As Object ' Outlook.Account
        
    '���łɋN�����Ă���Outlook�A�v���P�[�V�������Q�Ƃ���
    'Outlook���N�����Ă��Ȃ��ꍇ�͉����������̏����ɐi�ށB
    On Error Resume Next
    Set objOl = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    '�A�v�����N�����Ă��Ȃ��ꍇOutlook�A�v���P�[�V�����N��
    If objOl Is Nothing Then
        Set objOl = CreateObject("Outlook.Application")
    End If
    
    Set objMl = objOl.CreateItem(0) 'olMailItem

    With objMl
        
        .To = toAdd 'To�A�h���X
        .cc = ccAdd '
        
        For Each Account In objOl.Session.accounts
            If Account.smtpAddress = sendAccount Then
                Set tgtAcc = Account
                Exit For
            End If
        Next
        
        If tgtAcc Is Nothing Then
            GoTo err
        End If
        
        Set .SendUsingAccount = tgtAcc ' �A�J�E���g�w��
        .subject = subject '����
        .Body = msgBody '�{��
        On Error GoTo err
        .send '���[�����M
        On Error GoTo 0
    End With
    
    sendMail = True
    
nrmFin:
    Set objMl = Nothing
    Set objOl = Nothing

Exit Function
err:
    opeLog.Add "���M���G���[�ɂ��A���[�g���[���͑����܂���ł����B"
    GoTo nrmFin

End Function
