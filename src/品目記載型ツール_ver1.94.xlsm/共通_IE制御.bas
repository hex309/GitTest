Attribute VB_Name = "共通_IE制御"
Option Explicit

Public Sub ieCheckSSISLogin(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)
    
    Dim 検索ワード1 As String, 検索ワード2 As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1

        Dim oTAG As Object

        Dim errorMsgFlg As Boolean: errorMsgFlg = False
        Dim メッセージ As String

        For Each oTAG In oIE.document.getElementsByTagName("html")
            If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) <> 0 Then errorMsgFlg = True
            If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) <> 0 Then errorMsgFlg = True
            Err.Clear
            
            If errorMsgFlg = True Then
                メッセージ = "ログインに失敗しました。オートパイロットは停止しました。"
                MessageBox 0, メッセージ, "確認", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION
                End
            End If
        Next
continue1:

    Next
    On Error GoTo 0


End Sub

'-------------------------
'IE存在チェック
'-------------------------
Public Sub ieExistCheck()

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
        If oWin.Name = "Internet Explorer" Then
            MsgBox "InternetExplorerを全部落としてから、再度、実行してください。", vbExclamation
            End
        End If
    Next

End Sub

'-------------------------
'アクション後の表示完了待機
'-------------------------
Public Sub ieWaitCheck()
   
    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    
ReTry:
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin
            
        Dim timeOut As Date
        
        '完全にページが表示されるまで待機する
        timeOut = Now + TimeSerial(0, 0, 20)
         
'        Do While oIE.Busy = True Or (oIE.readyState < 4 And oIE.readyState > 0)
        Do While oIE.Busy = True Or (oIE.readyState < 4)
            If Err.Number <> 0 Then GoTo ReTry
'            DoEvents
            Sleep 250
            If Now > timeOut Then
                Debug.Print "【A】oIE.ReadyState:" & oIE.readyState & " hWnd:" & oWin.hWnd & " LocationName:" & oWin.LocationName
                oIE.Refresh
                timeOut = Now + TimeSerial(0, 0, 20)
            End If
        Loop
        
        Sleep 250
        timeOut = Now + TimeSerial(0, 0, 20)
        
        Do While oIE.document.readyState <> "complete"
            If Err.Number <> 0 Then GoTo ReTry
'            DoEvents
            Sleep 250
            If Now > timeOut Then
                Debug.Print "【B】oIE.ReadyState:" & oIE.readyState & " hWnd:" & oWin.hWnd & " LocationName:" & oWin.LocationName
                oIE.Refresh
                Sleep 1000
                timeOut = Now + TimeSerial(0, 0, 20)
            End If
        Loop
        On Error GoTo 0

continue1:
    Next

End Sub

'-------------------------
'アクション後の表示完了待機
'-------------------------
'Public Sub ieWaitCheck(ByVal oIE As Object)
'
'    Dim timeOut As Date
'
'    '完全にページが表示されるまで待機する
'    timeOut = Now + TimeSerial(0, 0, 20)
'
'    Debug.Print "oIE.ReadyState:" & oIE.readyState
'
'    Do While oIE.Busy = True Or (oIE.readyState < 4 And oIE.readyState > 0)
'        DoEvents
'        Sleep 250
'        If Now > timeOut Then
'            oIE.Refresh
'            timeOut = Now + TimeSerial(0, 0, 20)
'        End If
'    Loop
'
'    Sleep 250
'
'    timeOut = Now + TimeSerial(0, 0, 20)
'
'    Do While oIE.document.readyState <> "complete"
'        DoEvents
'        Sleep 1
'        If Now > timeOut Then
'            oIE.Refresh
'            timeOut = Now + TimeSerial(0, 0, 20)
'        End If
'    Loop
'
'End Sub
