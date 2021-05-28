VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorpSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCorpName As String
Private mSiteName As String
Private mBaseUrl As String
Private mCorpID As String
Private mUserName As String
Private mUserPass As String
Private mLastUpdate As Date
Private mDLLayout As String
Private mDLLayoutEV As String
Private mNavPDLayNo As Long
Private miWebPDLayNo As Long
Private miWebCrpID As String
Private WithEvents mObjIE As InternetExplorer
Attribute mObjIE.VB_VarHelpID = -1
Private ieTabs As Collection
Private defaultTabIdx As Long
Private isReady As Boolean
Private isError As Boolean
Private mMailFlg As Boolean
Private windowCloseFlg As Boolean

Private Sub Class_Initialize()
    Dim tmpIE As InternetExplorer
    
    On Error Resume Next
    Set tmpIE = New InternetExplorer
    On Error GoTo 0
    
    If tmpIE Is Nothing Then
        opeLog.Add "InternetExploreがクローズ処理中のため、起動できませんでした。"
    Else
        Set ieTabs = New Collection
        defaultTabIdx = 1
        
        ieTabs.Add tmpIE
        Set mObjIE = tmpIE
    End If

End Sub

Private Sub Class_Terminate()
    Dim ieTab As Variant
    
    If Not ieTabs Is Nothing Then
        If windowCloseFlg Then
            For Each ieTab In ieTabs
                On Error Resume Next
                ieTab.Quit
                Set ieTab = Nothing
                On Error GoTo 0
            Next
        End If
    End If
        
    Set ieTabs = Nothing
    Set ieTab = Nothing
    
End Sub

Public Property Get objIE(Optional ByVal tabIndex As Long = 0) As InternetExplorer
    If tabIndex = 0 Then
        tabIndex = defaultTabIdx
    End If
    
    If ieTabs Is Nothing Then
        Set objIE = Nothing
    Else
        Set mObjIE = ieTabs(tabIndex)
        Set objIE = ieTabs(tabIndex)
    End If

End Property

Public Property Get docIE(Optional ByVal tabIndex As Long = 0) As Object
    Set docIE = objIE(tabIndex).Document
End Property

Public Property Get baseURL() As String
    baseURL = mBaseUrl
End Property

Public Property Get siteName() As String
    siteName = mSiteName
End Property

Public Property Get corpName() As String
    corpName = mCorpName
End Property

Public Property Get CorpID() As String
    CorpID = mCorpID
End Property

Public Property Get userName() As String
    userName = mUserName
End Property

Public Property Get userPass() As String
    userPass = mUserPass
End Property

Public Property Get lastUpdate() As Date
    lastUpdate = mLastUpdate
End Property

Public Property Get dlLayout() As String
    dlLayout = mDLLayout
End Property

Public Property Get dlLayoutEV() As String
    dlLayoutEV = mDLLayoutEV
End Property

Public Property Get NavPDLayNo() As Long
    NavPDLayNo = mNavPDLayNo
End Property

Public Property Get iWebPDLayNo() As Long
    iWebPDLayNo = miWebPDLayNo
End Property

Public Property Get MailFlg() As Long
    MailFlg = mMailFlg
End Property

Public Property Get byClass(ByVal classValue As String, Optional ByVal tabIndex As Long = 0, Optional ByVal logFlg As Boolean = True) As Object
    Set byClass = getElement("class", classValue, tabIndex, logFlg)
End Property

Public Property Get byName(ByVal elementName As String, Optional ByVal tabIndex As Long = 0, Optional ByVal logFlg As Boolean = True) As Object
    Set byName = getElement("name", elementName, tabIndex, logFlg)
End Property

Public Property Get byTag(ByVal tagName As String, Optional ByVal tabIndex As Long = 0, Optional ByVal logFlg As Boolean = True) As Object
    Set byTag = getElement("tag", tagName, tabIndex, logFlg)
End Property

Public Property Get byId(ByVal idValue As String, Optional ByVal tabIndex As Long = 0, Optional ByVal logFlg As Boolean = True) As Object
    Set byId = getElement("id", idValue, tabIndex, logFlg)
End Property

Public Function setDefaultIETab(Optional ByVal tabIndex As Long = 1) As Boolean
    defaultTabIdx = tabIndex
    Set mObjIE = ieTabs(defaultTabIdx)
End Function

Public Function submit(ByRef tgtElement As Object, Optional ByVal timeOut As Date = #12:01:00 AM#, Optional ByVal wateState As Long = 4) As Boolean

    On Error GoTo err
    tgtElement.submit
    On Error GoTo 0
    
    If Not isLoad(timeOut, wateState) Then
        GoTo err
    End If
    
    submit = True
    
    Exit Function
err:
    If tgtElement Is Nothing Then
        opeLog.Add "操作対象の項目が見つかりませんでした。" & vbCrLf & "途中でエラーページに誘導されているか、Webのレイアウトが変更されていないか確認してください。"
    Else
        opeLog.Add "操作対象 : " & Left(tgtElement.innerText, 10) & vbCrLf & "操作対象に誤りがないかご確認ください。"
    End If
    
    submit = False

End Function

Public Function click(ByRef tgtElement As Object, Optional ByVal timeOut As Date = #12:01:00 AM#, Optional ByVal wateState As Long = 4) As Boolean
    
    On Error GoTo err
    tgtElement.click
    On Error GoTo 0
    
    If Not isLoad(timeOut, wateState) Then
        GoTo err
    End If
    
    click = True
    
    Exit Function
err:
    If tgtElement Is Nothing Then
        opeLog.Add "操作対象の項目が見つかりませんでした。" & vbCrLf & "途中でエラーページに誘導されているか、Webのレイアウトが変更されていないか確認してください。"
    Else
        opeLog.Add "操作対象 : " & Left(tgtElement.innerText, 10) & vbCrLf & "操作対象に誤りがないかご確認ください。"
    End If
    
    click = False
    
End Function

Public Function execScript(ByVal argScript As String, Optional ByVal wateState As Long = 4) As Boolean
    Me.docIE.parentWindow.execScript argScript
    
    If Not isLoad(, wateState) Then
        GoTo err
    End If
    
    execScript = True
Exit Function
err:
    execScript = False
    opeLog.Add "対象スクリプト : " & argScript & vbCrLf & "スクリプト実行に失敗したか、想定外の動作です。"

End Function

Public Function getTableCell(ByVal uniqValue As String, ByVal tgtCellValue As String, Optional offset As Long) As Object
    Dim table As Variant
    Dim i As Long
    
    For Each table In byTag("table")
        If InStr(table.innerText, uniqValue) > 0 Then
            For i = 0 To table.Cells.Length - 1
                If Trim(table.Cells(i).innerText) = tgtCellValue Then
                    Set getTableCell = table.Cells(i + offset)
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Function getElement(ByVal tgtElementName As String, ByVal tgtElementValue As String, Optional ByVal tabIndex As Long = 0, Optional ByVal logFlg As Boolean = True) As Object
    Dim tgtDoc As Object
    Dim i As Long
    
    Set tgtDoc = Me.docIE(tabIndex)
    tgtElementName = StrConv(tgtElementName, vbLowerCase)
    
    Do
        On Error Resume Next
        Select Case tgtElementName
            Case "id"
                Set getElement = tgtDoc.getElementById(tgtElementValue)
            Case "tag"
                Set getElement = tgtDoc.getElementsByTagName(tgtElementValue)
            Case "class"
                Set getElement = tgtDoc.getElementsByClassName(tgtElementValue)
            Case "name"
                Set getElement = tgtDoc.getElementsByName(tgtElementValue)
            Case Else
                'Do nothing
        End Select
        On Error GoTo 0
        
        If Not getElement Is Nothing Then
            Exit Do
        Else
            DoEvents
            Application.Wait [now() + "00:00:00.5"]
        End If
        
        i = i + 1
        
    Loop Until i = 10
    
    If getElement Is Nothing And logFlg Then
        opeLog.Add tgtElementValue & "という値の" & tgtElementName & "タグは見つかりませんでした。"
    End If

End Function

Public Function navigate(ByVal TargetURL As String, Optional ByVal tabIndex As Long = 0, Optional ByVal ViewFlg As Boolean = True) As Boolean
   
    On Error GoTo err
    Me.objIE(tabIndex).navigate TargetURL
    On Error GoTo 0
    
    If Not isLoad Then
        GoTo err
    End If
    
    mObjIE.visible = ViewFlg
    navigate = True
Exit Function
err:
    navigate = False
    opeLog.Add "対象URL : " & TargetURL & vbCrLf & "URLに誤りがないかご確認ください。"

End Function

Public Function Refresh(Optional ByVal tabIndex As Long = 0) As Boolean
    Me.objIE(tabIndex).Refresh
        
    If Not isLoad Then
        Refresh = False
        opeLog.Add "対象URL:" & mObjIE.LocationURL & vbCrLf & "ページのリフレッシュに失敗しました。"
        Exit Function
    End If
    
    Refresh = True

End Function

Public Function post(ByVal TargetURL As String, _
                   Optional ByVal tabIndex As Long = 0, _
                   Optional ByVal PostData As String = vbNullString, _
                   Optional ByVal ViewFlg As Boolean = True, _
                   Optional ByVal Headers As String = "Content-Type: application/x-www-form-urlencoded" & vbCrLf, _
                   Optional ByVal Flags As Long = 0, _
                   Optional ByVal TargetFrame As String = vbNullString) As Boolean
                   
    Dim bPostData() As Byte
    
    'PostDataをバイナリ化
    bPostData = StrConv(PostData, vbFromUnicode)
    
    '指定したURLのページを表示する
    Me.objIE(tabIndex).navigate TargetURL, Flags, TargetFrame, bPostData, Headers
 
    'IE(InternetExplorer)が完全表示されるまで待機
    If Not isLoad Then
        post = False
        opeLog.Add "対象URL : " & TargetURL & vbCrLf & "URLに誤りがないかご確認ください。"
        Exit Function
    End If
    
    'IE(InternetExplorer)を表示・非表示
    mObjIE.visible = ViewFlg
    
    post = True
    
End Function

Private Function isLoad(Optional ByVal timeOut As Date = #12:01:00 AM#, Optional ByVal wateState As Long = 4) As Boolean
    '完全にページが表示されるまで待機する
    timeOut = Now + timeOut
    
    Do Until isReady
        DoEvents
        
        If isError Then
            opeLog.Add "ネットワーク側の要因により開けませんでした。"
            isError = False
            Exit Function
        ElseIf Now > timeOut Then
            opeLog.Add "接続中にタイムアウトしました。"
            Exit Function
        ElseIf cancelFlg Then
             opeLog.Add "キャンセルされました。"
            Exit Function
        End If
        
        'refresh 時にDocument Completeイベントが検知されないので
        If mObjIE.ReadyState = wateState And mObjIE.Busy = False Then
            If mObjIE.Document.ReadyState = "complete" Then
                isReady = True
            End If
        End If
    Loop
    
    isReady = False
    isLoad = True
    
End Function

Public Sub quitIE(Optional ByVal tabIndex As Long = 0)
    If tabIndex = 0 Then
        tabIndex = defaultTabIdx
    End If
    
    ieTabs(tabIndex).Quit
    ieTabs.Remove tabIndex
    
    defaultTabIdx = ieTabs.Count
    Set mObjIE = ieTabs(defaultTabIdx)
    
End Sub

Public Sub quitAll()
    windowCloseFlg = True
End Sub

Public Sub visible(ByVal argVisible As Boolean)
    Dim ieTab As Variant
    
    For Each ieTab In ieTabs
        On Error Resume Next
        ieTab.visible = argVisible
        On Error GoTo 0
    Next
End Sub

'変数と一致するURLがIE上で開いている場合、IEを閉じる
Public Sub cleanUpTgtSite()
    Dim objShell As Object
    Dim objWin As Object

    Set objShell = CreateObject("Shell.Application")
    For Each objWin In objShell.Windows
        If objWin.name = "Internet Explorer" Then
            If InStr(objWin.LocationURL, mBaseUrl) > 0 Then
                objWin.Quit  '対象IEを閉じる
            End If
        End If
    Next

End Sub

Public Function getFileNameNow(ByVal argSiteName As String, ByVal tgtDataType As Long) As String
    Dim tgtDataName As String
    Dim tgtSiteName As String
    Dim i As Long

    Select Case argSiteName
        Case "リクナビ"
            tgtSiteName = "リク"
        Case "マイナビ"
            tgtSiteName = "マイ"
        Case Else
            tgtSiteName = "他"
    End Select
    
    Select Case tgtDataType
        Case dataType.personal
            tgtDataName = "_個人"
        Case dataType.Seminar
            tgtDataName = "_セミナ"
        Case Else
            tgtDataName = "_"
    End Select
    
    getFileNameNow = miWebCrpID & mCorpName & "_" & tgtSiteName & tgtDataName & Format(Date, "mmdd") & Format(Now, "hhmmss")
    
    If argSiteName = "マイナビ" Then
        getFileNameNow = chopString(getFileNameNow, 40)
    ElseIf argSiteName = "リクナビ" Then
        getFileNameNow = chopString(getFileNameNow, 50)
    End If
    
    i = 0
    
    'ファイル存否チェック。末尾番号で99以上重複すると想定外。v1.09現在IDが30桁とかに増えない限り起きない想定なので考慮しません。ごめん。
    Do Until getDlFilePath(getFileNameNow & ".txt", False) = vbNullString Or i = 99
        getFileNameNow = Left(getFileNameNow, Len(getFileNameNow) - 3) & "_" & Format(i, "00")
        i = i + 1
    Loop
        
End Function

Public Function waitPageMoved(ByVal tgtURL As String) As Boolean
    Dim timeOut As Date

    'ページの遷移完了まで待機する
    timeOut = Now + TimeSerial(0, 3, 0)
    
    Do While True
        DoEvents
        
        If Now > timeOut Then
            opeLog.Add "入力待ちでタイムアウトしました。"
            Exit Do
        ElseIf cancelFlg Then
            opeLog.Add "キャンセルされました。"
            Exit Do
        End If
    
        If InStr(mObjIE.LocationURL, tgtURL) > 0 Then
            waitPageMoved = True
            Exit Do
        End If
    Loop

End Function

'「アカウント」シート表のデータを取得、データの行ごとにループ処理
'引数は「アカウント」シートの対象企業名かつ対象サイト名
Public Function setCorp(ByVal tgtCorpName As String, ByVal tgtSiteName As String, Optional ByVal loginOnly As Boolean = False) As Boolean
    Dim AccData As Variant
    Dim tgtRng As Range
    Dim i As Long, j As Long
    
    '
    mCorpName = tgtCorpName
    mSiteName = tgtSiteName
    
    Set tgtRng = getCurrentRegion(AccountSh.Cells(1, 2), 1, False)
    
    If tgtRng Is Nothing Then
        opeLog.Add "アカウントが記載されていません！" & vbCrLf _
             & AccountSh.name & "シートをご確認ください。"
        Exit Function
    Else
        AccData = tgtRng.Value
    End If
    
    '表の値に対し、余分なsペースの削除および、「なし」または「無し」の場合は空欄を取得
    For i = LBound(AccData, 1) To UBound(AccData, 1)
        For j = LBound(AccData, 2) To UBound(AccData, 2)
            If Trim(AccData(i, j)) = "なし" Or Trim(AccData(i, j)) = "無し" Then AccData(i, j) = vbNullString
        Next
    Next
    
    For i = LBound(AccData, 1) To UBound(AccData, 1)
        '引数と一致する、対象企業名かつ対象サイト名の場合、同じ行のデータを取得
        If AccData(i, AC_CORP_IDX) = mCorpName Then
            If AccData(i, AC_SITE_IDX) = mSiteName Then
            
                mBaseUrl = AccData(i, AC_ADDR_IDX)
                mCorpID = AccData(i, AC_COID_IDX)
                mUserName = AccData(i, AC_ACNT_IDX)
                mUserPass = AccData(i, AC_PSWD_IDX)
                mDLLayout = AccData(i, AC_DLLO_IDX)
                mDLLayoutEV = AccData(i, AC_EDLO_IDX)
                On Error Resume Next
                mNavPDLayNo = AccData(i, AC_IWUL_IDX)
                miWebPDLayNo = AccData(i, AC_IWDL_IDX)  '個人用ダウンロードレイアウトインデックスNo
                On Error GoTo 0
                miWebCrpID = AccData(i, AC_IWCO_IDX)
                
                setCorp = True
                Exit For
            End If
        End If
    Next
    
    If Not setCorp Then
        opeLog.Add mSiteName & "の" & tgtCorpName & "様用のアカウント、パスワードが見つかりません。" & vbCrLf _
             & AccountSh.name & "シートをご確認ください。"
        Exit Function
    End If
    
    setCorp = False
       
    Set tgtRng = getCurrentRegion(ScenarioSh.Cells(1, 1), 1, False)
    
    If Not tgtRng Is Nothing Then
        AccData = tgtRng.Value
        '「メイン」シートの表データの行数分(項目名「項番」が10まであるため、基本10行)ループ処理
        For i = LBound(AccData, 1) To UBound(AccData, 1)
            If AccData(i, SC_CORP_IDX) = mCorpName Then
                '引数の対象企業名と、「メイン」シートの項目名「対象企業名」データが一致した場合
                On Error Resume Next
                '最終更新日を登録
                mLastUpdate = CDate(AccData(i, SC_LAST_UPDT_COL_IDX))
                On Error GoTo 0
                mMailFlg = AccData(i, SC_MLFL_IDX)
                '「メイン」シートの項目名「最終完了日時」にデータがない場合
                If mLastUpdate = 0 Then
                    mLastUpdate = #6/1/2018#  '日付を取得
                End If
        
                setCorp = True
                Exit For
            End If
        Next
    End If
        
    If loginOnly Then
        setCorp = True
    ElseIf Not setCorp Then
        opeLog.Add tgtCorpName & "様の実行指定が見つかりません。" & vbCrLf _
             & ScenarioSh.name & "シートをご確認ください。"
        Exit Function
    End If

End Function

Private Sub mObjIE_NewWindow2(ppDisp As Object, Cancel As Boolean)
    Dim newTab As InternetExplorer
        
    Set newTab = New InternetExplorer
    Set ppDisp = newTab
    
    ieTabs.Add newTab
    defaultTabIdx = ieTabs.Count
    
    Set mObjIE = newTab
    
End Sub

Private Sub mObjIE_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    isReady = False
    isError = False
End Sub

Private Sub mObjIE_NavigateError(ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    If mObjIE.LocationURL = URL Or mObjIE.LocationURL = vbNullString Then
        isError = True
        Cancel = True
    End If
End Sub

Private Sub mObjIE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If mObjIE.LocationURL = URL Then
        isReady = True
    End If
End Sub

Private Sub mObjIE_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
    Dim tabIndex As Long
        
    tabIndex = defaultTabIdx
    
    'ieTabs(tabIndex).Quit
    ieTabs.Remove tabIndex
    
    defaultTabIdx = ieTabs.Count
    Set mObjIE = ieTabs(defaultTabIdx)
End Sub

