Attribute VB_Name = "MyNaviOpe"
Option Explicit
Option Private Module

Public Function loginMyNavi(ByRef argTgtSite As CorpSite) As Boolean
    With argTgtSite
        On Error GoTo loginErr
        If Not .navigate(.baseURL, , True) Then Exit Function
        
        .byId("mwCorpNo").Value = .CorpID
        .byId("empStaffPasswd").Value = .userPass
        .byId("tmpEmpStaffId").Value = .userName
        If Not .click(.byId("doLogin")) Then
            GoTo loginErr
        End If
        
        If .byId("errorsArea", , False) Is Nothing Then
            loginMyNavi = True
        Else
            opeLog.Add Trim(.byId("errorsArea").innerText)
            GoTo loginErr
        End If
        
        On Error GoTo 0
    End With
    

Exit Function

loginErr:
    opeLog.Add "マイナビにログインできませんでした。" & vbCrLf & "アドレス/アカウント/パスワード/通信状況をご確認ください。"
    loginMyNavi = False
        
End Function

Public Function moveSearchWindow(ByRef argTgtSite As CorpSite, Optional ByVal fromTop As Boolean = True) As Boolean
    With argTgtSite
        If fromTop Then
            'comMiwsTopLink押下後も正常な画面に遷移せず、searchTopが押せない状況がたまに発生するとのこと。
            '原因不明だが、待機を入れて様子見（2019/4/19）
            Sleep 2000
            
            .click .byId("comMiwsTopLink")
            .click .byId("searchTop")
        Else
            .click .byId("searchTop")
        End If
    End With
    
    moveSearchWindow = True
End Function

Public Function searchMyNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As Boolean
    Dim ancElmt As Object
    Dim altMsg As String

    With argTgtSite
        If tgtDtType = dataType.personal Then
            If Not .click(.byId("searchDeliveryList")) Then Exit Function
            
            .byId("regdateBeforeYear").Value = Format(.lastUpdate, "yyyy")
            .byId("regdateBeforeMonth").Value = Format(.lastUpdate, "mm")
            .byId("regdateBeforeDay").Value = Format(.lastUpdate, "dd")
            .byId("regdateCond").Value = 2
            .byId("regdateAfterYear").Value = Format(Date, "yyyy")
            .byId("regdateAfterMonth").Value = Format(Date, "mm")
            .byId("regdateAfterDay").Value = Format(Date, "dd")
            
            For Each ancElmt In .byId("main").getElementsByClassName("actbtn")
                If ancElmt.id = "doRegist" Then
                    If Not .click(ancElmt) Then Exit Function
                    Exit For
                End If
            Next
            
             altMsg = Format(.lastUpdate, "yyyy/mm/dd") & " ～ " & Format(Date, "yyyy/mm/dd") & "で検索しましたが該当するデータはありませんでした。"
            
        ElseIf tgtDtType = dataType.Seminar Then
            'Do Nothing 全件検索
            
             altMsg = "いずれかの日程に予約のあるデータはありませんでした。"
        Else
            Exit Function
        End If
        
        If Not .click(.byId("doSearch")) Then Exit Function
        
        If InStr(.byId("main").innerText, "検索結果は0件でした") > 0 Then
            opeLog.Add altMsg
            Exit Function
        End If
        
        .byId("executeUpList").Value = "c_file_out_reserve_edit"
        If Not .click(.byId("doExecuteUp")) Then Exit Function
        
    End With
    
    searchMyNaviDt = True
    
End Function

Public Function moveSearchWindowOld(ByRef argTgtSite As CorpSite) As Boolean
    With argTgtSite

        '#リンクをたどるパターン
        .click .byId("comMiwsTopLink")
        .click .byId("searchTop")
        .click .byId("searchDeliveryList")
    End With
    
    moveSearchWindowOld = True

End Function

Public Function searchMyNaviPd(ByRef argTgtSite As CorpSite) As Boolean
    
    With argTgtSite
        .byId("regdateBeforeYear").Value = Format(.lastUpdate, "yyyy")
        .byId("regdateBeforeMonth").Value = Format(.lastUpdate, "mm")
        .byId("regdateBeforeDay").Value = Format(.lastUpdate, "dd")
        .byId("regdateCond").Value = 2
        .byId("regdateAfterYear").Value = Format(Date, "yyyy")
        .byId("regdateAfterMonth").Value = Format(Date, "mm")
        .byId("regdateAfterDay").Value = Format(Date, "dd")
        
        Dim ancElmt As Variant
    
        For Each ancElmt In .byId("main").getElementsByClassName("actbtn")
            If ancElmt.id = "doRegist" Then
                ancElmt.click
                Exit For
            End If
        Next
        
        IECheck .objIE
        
        .click .byId("doSearch")

        If InStr(.byId("main").innerText, "検索結果は0件でした") > 0 Then
            opeLog.Add Format(.lastUpdate, "yyyy/mm/dd") & " ～ " & Format(Date, "yyyy/mm/dd") & "で検索しましたが該当するデータはありませんでした。"
            Exit Function
        End If
        
        .docIE.getElementById("executeUpList").Value = "c_file_out_reserve_edit"
        .docIE.getElementById("doExecuteUp").click
        IECheck .objIE
        
    End With
    
    searchMyNaviPd = True

End Function

Public Function makeMyNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As String
    Dim table As Variant
    Dim dlLayout
    Dim fileName As String
    Dim tallyho As Boolean
    Dim i As Long
    
    With argTgtSite
        
        If tgtDtType = dataType.personal Then
            dlLayout = .dlLayout
        ElseIf tgtDtType = dataType.Seminar Then
            dlLayout = .dlLayoutEV
        Else
            Exit Function
        End If
    
        For Each table In .byTag("table")
            If table.Cells(0).innerText = "選択" Then
                For i = 0 To table.Cells.Length - 1
                    If table.Cells(i).innerText = dlLayout Then
                        table.Cells(i - 1).Children(0).click
                        tallyho = True
                        Exit For
                    End If
                Next
                Exit For
            ElseIf InStr(table.innerText, "エラー") > 0 Then
                opeLog.Add Replace(.byId("main").innerText, vbCrLf, vbNullString)
                Exit Function
            End If
        Next
        
        If Not tallyho Then
            opeLog.Add "『" & dlLayout & "』といったマイナビレイアウト名はありません。"
            Exit Function
        End If
        
        fileName = .getFileNameNow("マイナビ", tgtDtType)
        
        .docIE.getElementById("name").Value = fileName
        .docIE.getElementById("downloadFormat[v_10]").click
        
        If Not .click(.byId("doConfirm")) Then Exit Function
        If Not .click(.byId("doRegist")) Then Exit Function
        
        .quitIE
        
        If Not .click(.byId("iOOutExe")) Then Exit Function
        
    End With
    
    makeMyNaviDt = fileName
    
End Function

Public Function chkDateCreated(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As Boolean
    Dim timeOut As Date

    timeOut = Now + SettingSh.DlTimeOut
    
    With argTgtSite
        Do While InStr(.getTableCell("選択", argFileName, 3).innerText, "作成中") > 0
            .Refresh
            
            If Now > timeOut Then
                opeLog.Add "データ出力中ににタイムアウトしました。"
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "キャンセルされました。"
                Exit Function
            End If
            
            Sleep 5000
            DoEvents
        Loop
    End With
    
    chkDateCreated = True
End Function

Public Function dlMyNaviCSV(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As String
    Dim timeOut As Date
    Dim filePath As String
    Dim alrPass As Boolean
    
    timeOut = Now + SettingSh.DlTimeOut
    
    With argTgtSite
        .getTableCell("選択", argFileName, -1).Children(0).click
        
        If .byId("authId").Value <> vbNullString Then
            .click .byId("doOutput"), , 3
            
            If .byClass("midashi_1")(0).innerText = "エラー" Then
                If Not .click(.byId("iOOutExe")) Then Exit Function
            Else
                alrPass = True
            End If
        End If
        
        If Not alrPass Then
            .byId("authId").Focus
            
            AlertBox.Label1.ForeColor = &HFF&
            AlertBox.Label1 = "文字認証を入力してボタンを押してください！"
            
            sendCaptAlert "マイナビのダウンロード認証待ちです。"
        End If
        
        Do
            DoEvents
            
            On Error Resume Next
            pushSaveButton .objIE
            On Error GoTo 0
        
            'Application.Wait [now() + "00:00:00.5"]
            filePath = getDlFilePath(argFileName & ".txt", False)
            
            If Now > timeOut Then
                getDlFilePath (argFileName & ".txt")
'                'タイムアウトのメッセージは呼び出し先で発報
'                AlertBox.Label1.ForeColor = &H80000012
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "キャンセルされました。"
                Exit Function
            End If
            
        Loop While filePath = vbNullString
        
    End With
    
'    AlertBox.Label1.ForeColor = &H80000012
    dlMyNaviCSV = filePath
        
End Function

Public Function getDiffFile(ByVal mynavSmDataFilePath As String, ByVal tgtCorpName As String, ByVal lastUpdate As Date) As String

    ' 前回のセミナーファイルパスの行番号
    Dim oldPathRow As Long
    
    oldPathRow = SettingSh.OldMyNaviRowIndex(tgtCorpName)
    
    If oldPathRow = 0 Then
        opeLog.Add SettingSh.name & "シートに対象企業名『" & tgtCorpName & "』がありません。"
        Exit Function
    End If

    ' 前回のセミナーファイルパス存在チェック用
    Dim CheckPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CheckPath = SettingSh.Cells(oldPathRow, 2).Value
            
    'パスが空白か、記載されていてかつファイルがあるときは、差分ファイル生成
    '空白の場合は差分ファイル生成時に日付だけlastUpdateに更新される。
    If CheckPath = vbNullString Or fso.FileExists(CheckPath) Then
        getDiffFile = makeDiffFile(mynavSmDataFilePath, CheckPath, lastUpdate)
        
        If getDiffFile = vbNullString Then Exit Function
        
        SettingSh.Cells(oldPathRow, 3).Value = mynavSmDataFilePath
    Else
    'パスの記載があり、ファイルが無い場合はエラー
        opeLog.Add SettingSh.name & "の[B" & oldPathRow & "]に記載されたパスにファイルが存在しません"
    End If
    
End Function

Public Function getMynaviCancelDate(ByRef argTgtSite As CorpSite, _
                                    ByVal myNaviId As String, _
                                    ByVal semID As String, _
                                    Optional ByVal tabIndex As Long) As Date

    With argTgtSite
        .byId("searchCode", tabIndex).Value = myNaviId
        .click .byId("doSearchMenu", tabIndex)
        .click .byId("entryLink", tabIndex)
        
        IECheck .objIE
        
        Dim linkTag As Object
        
        For Each linkTag In .byId("linktab").getElementsByTagName("a")
            If InStr(linkTag.innerText, "説明会・面接") > 0 Then
                .click linkTag
                Exit For
            End If
        Next
        
        Dim table As Variant
        Dim update As Date
        Dim cancelDate As Date
        Dim tempDate As Date
        Dim i As Long
        
        For Each table In .byTag("table")
            If InStr(table.innerText, "説明会・面接予約状況一覧") > 0 Then
                For i = 0 To table.Cells.Length - 1
                    If Trim(table.Cells(i).innerText) = semID Then
                        tempDate = CDatePlus(Trim(table.Cells(i + 5).innerText))
                        If tempDate >= update Then
                            update = tempDate
                            cancelDate = CDatePlus(Trim(table.Cells(i - 1).innerText))
                        End If
                    End If
                Next
                Exit For
            End If
        Next
        
    End With
    
    getMynaviCancelDate = cancelDate
    
End Function

Public Function CDatePlus(ByVal argStr As String) As Date
    Dim buf As String
    
    buf = Trim(argStr)
    buf = Replace(buf, vbCr, vbNullString)
    buf = Replace(buf, vbLf, vbNullString)
    
    If buf = vbNullString Then
        CDatePlus = 0
    Else
        CDatePlus = CDate(buf)
    End If
    
End Function
