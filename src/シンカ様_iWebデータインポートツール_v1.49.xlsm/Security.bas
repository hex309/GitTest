Attribute VB_Name = "Security"
Option Explicit
Option Private Module

'「メイン」シートのボタン名を、「保護解除」から「3分間保護解除」へ変更
'右横のメッセージを「＊解除中はデータインポートは実行できません」にする
Public Function ReProtect(ByVal sheetName As String, ByVal endTime As Date)
    
    On Error Resume Next
    '引数(endTime)の時刻になったら、当プロシージャの処理を実行
    Application.OnTime endTime, "'ReProtect""" & Worksheets(sheetName).name & """,""" & endTime & """ '", , False
    On Error GoTo 0  'エラーを無効にする
    
    'ボタン名を「3分間保護解除」へ変更
    '右横のメッセージを「＊解除中はデータインポートは実行できません」にする
    Worksheets(sheetName).protectSheet
    
    Application.EnableEvents = True

    MsgBox Worksheets(sheetName).name & "シートの保護を再開しました。", vbInformation

End Function

'シート保護を解除しPasswordが正しく入力された場合、ボタン横にフォント色赤で「保護解除」終了時間を明示
'ボタン名を「3分間保護解除」から「保護再開」へ変更
Public Sub unprotectFewMinutes(ByVal sheetName As String, ByVal endTime As Date, Optional ByVal dspMsgAdd As String)
    Dim sh As Shape

    On Error GoTo Error  'Passwordが違った場合
    Worksheets(sheetName).Unprotect  'シート保護を解除するため、Password入力画面を表示
    
    On Error GoTo 0

    'Password入力画面で、「×」または「キャンセル」を押下された場合は、処理終了
    If Worksheets(sheetName).ProtectContents Then
        Exit Sub
    End If

    '引数(endTime)の時間になったら、保護再開処理を行う、Functionプロシージャ「ReProtect」(直上)を実行
    Application.OnTime endTime, "'ReProtect""" & Worksheets(sheetName).name & """,""" & endTime & """ '", , True
    Application.EnableEvents = False
    
    If dspMsgAdd <> vbNullString Then  '引数(dspMsgAdd)に値がある場合は、値にフォント色赤で書換
        With Worksheets(sheetName).Range(dspMsgAdd)
            .Value = "* " & Format(endTime, "hh時mm分ss秒") & " に保護を再開します。"
            .Font.Color = vbRed
        End With
    End If

    For Each sh In Worksheets(sheetName).Shapes  '保護ボタンの場合、ボタン名を書換
        If sh.name = "ProtectBtn" Then
             sh.TextEffect.Text = "保護再開"
        End If
    Next

    MsgBox "保護を解除しました。" & vbCrLf & Format(endTime, "hh時mm分ss秒") & " に保護を再開します", vbExclamation

Exit Sub
Error:
    MsgBox "パスワードが違います。"

End Sub

Public Function sendCaptAlert(ByVal msgBody As String) As Boolean
    Dim userAcc As String
    Dim sndAdd As String
    Dim msgSubject As String
    
    If ScenarioSh.getUserName = vbNullString Then Exit Function
    
    With MailSettingSh
        userAcc = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = .getSendAccount(ScenarioSh.getUserName)
        sndAdd = sndAdd & "; " & .getSendAccount(, "認証/追加")
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
        sndAdd = sndAdd & "; " & .getSendAccount(, "認証/追加")
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
        msgSubject = IIf(tgtCorp = vbNullString, vbNullString, "【" & tgtCorp & "】") & .getFinSubject
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
        
    'すでに起動しているOutlookアプリケーションを参照する
    'Outlookが起動していない場合は何もせず次の処理に進む。
    On Error Resume Next
    Set objOl = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    'アプリが起動していない場合Outlookアプリケーション起動
    If objOl Is Nothing Then
        Set objOl = CreateObject("Outlook.Application")
    End If
    
    Set objMl = objOl.CreateItem(0) 'olMailItem

    With objMl
        
        .To = toAdd 'Toアドレス
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
        
        Set .SendUsingAccount = tgtAcc ' アカウント指定
        .subject = subject '件名
        .Body = msgBody '本文
        On Error GoTo err
        .send 'メール送信
        On Error GoTo 0
    End With
    
    sendMail = True
    
nrmFin:
    Set objMl = Nothing
    Set objOl = Nothing

Exit Function
err:
    opeLog.Add "送信時エラーによりアラートメールは送られませんでした。"
    GoTo nrmFin

End Function
