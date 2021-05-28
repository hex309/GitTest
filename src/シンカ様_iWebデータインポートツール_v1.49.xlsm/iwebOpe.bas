Attribute VB_Name = "iwebOpe"
Option Explicit
Option Private Module

Public Function loginiWeb(argIWeb As CorpSite) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    
    tgtURL = "login/check"
    
    With argIWeb
        pstData = "id=" & .userName & "&pass=" & .userPass
        If Not .post(.baseURL & tgtURL, , pstData, True) Then GoTo loginErr 'ログイン
        If InStr(.objIE.LocationURL, "top/index") = 0 Then GoTo loginErr
        
        
'        If InStr(.byId("front").innerText, "tabcd = ""all"";") = 0 Then
'            .click .byId("taball")
'            Do While InStr(.byId("front").innerText, "tabcd = ""all"";") = 0
'                Sleep 1000
'            Loop
'        End If
    End With
    
    loginiWeb = True
    
Exit Function
loginErr:
    opeLog.Add "iWebにログインできませんでした。" & vbCrLf & "アドレス/アカウント/パスワード/通信状況をご確認ください。"
    loginiWeb = False

End Function

Public Function selectTab(argIWeb As CorpSite, ByVal jobTabName As String) As Boolean
    Dim hitFlg As Boolean
    Dim aTag As Object

    With argIWeb
        If .byId("crnt").innerText = jobTabName Then
            selectTab = True
            Exit Function
        End If
        
        For Each aTag In .byId("tabArea").getElementsByTagName("a")
            If aTag.innerText = jobTabName Then
                .click aTag
                hitFlg = True
                Exit For
            End If
        Next
        
        If Not hitFlg Then Exit Function
    
        Do While .byId("loadArea").innerText = "now loading..."
            Sleep 100
        Loop
    End With
    
    selectTab = True
         
End Function

Public Function DLiWebData(argIWeb As CorpSite, ByVal dataTypeCode As Long) As String
    Dim objIE As Object 'InternetExplorer
    Dim tgtURL As String
    Dim pstData As String
    
    opeLog.Add "iWebにログイン開始.. アカウント：" & argIWeb.userName
    If Not loginiWeb(argIWeb) Then Exit Function
    opeLog.Add "iWebにログイン完了。個人情報ダウンロードページへ移動開始.."
    
    With argIWeb
    
        tgtURL = "download/confirm/"
        pstData = "layflg=cd&dlcd=" & dataTypeCode & "&ptflg=all"
        
        If Not .post(.baseURL & tgtURL, , pstData) Then Exit Function
        opeLog.Add "移動完了。CSVファイル名確認中..."
        
        Dim Text As Variant
        Dim fileName As String
    
        For Each Text In Split(.byId("selectTable").innerText, vbCrLf)
            If InStr(Text, "csv") > 0 Then
                fileName = Trim(Text)
                Exit For
            End If
        Next
    
        If fileName = vbNullString Then
            opeLog.Add "ダウンロードするCSVのファイル名が確認できませんでした。"
            Exit Function
        End If
        
        opeLog.Add "CSVファイル名確認。ダウンロード開始..."
        .click .byId("btnDownload"), SettingSh.DlTimeOut, 3
    
        Dim timeOut As Date
        
        timeOut = Now + SettingSh.DlTimeOut
    
        Do
            On Error Resume Next
            pushSaveButton .objIE
            On Error GoTo 0
        
            Application.Wait Now + TimeValue("00:00:01")
            DLiWebData = getDlFilePath(fileName, False)
            
            If Now > timeOut Then
                getDlFilePath (fileName)
                'タイムアウトのメッセージは呼び出し先で発報
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "キャンセルされました。"
                Exit Function
            End If
        
        Loop While DLiWebData = vbNullString
    End With
    
    opeLog.Add "ダウンロード完了。"
    
End Function

Private Function loadNayoseTb(ByVal argTable As Object) As Object
    Dim cell As Variant

    If argTable.className = "eventTableWidth700" Then
    
        Set loadNayoseTb = CreateObject("Scripting.Dictionary")
            
        For Each cell In argTable.Cells
        
            If cell.tagName = "TH" Then
                loadNayoseTb.Add cell.innerText, cell.nextElementSibling.innerText
                
            ElseIf cell.cellIndex = 0 Then
                loadNayoseTb.Add "選択", cell
            ElseIf InStr(cell.innerText, "大学") > 0 Or InStr(cell.innerText, "学校") Then
                loadNayoseTb.Add "学校", cell.innerText
            ElseIf InStr(cell.innerText, "HMI") > 0 And Len(cell.innerText) = 10 Then
                loadNayoseTb.Add "ID", cell.innerText
            End If
        Next
    End If

End Function

Private Function isAllOvreWrite(ByVal nyTable As Object, ByVal navTable As Object) As Boolean
    Dim cell As Variant
    
    isAllOvreWrite = True
    
    For Each cell In nyTable
        If cell <> "ID" And nyTable(cell) <> navTable(cell) Then
            opeLog.Add IIf(Trim(nyTable(cell)) = vbNullString, "空白", "(i-Web)：" & nyTable(cell)) & _
                        "← (ナビ)：" & IIf(Trim(navTable(cell)) = vbNullString, "空白", navTable(cell))
            If nyTable(cell) = " " Then
                isAllOvreWrite = isAllOvreWrite And True
            ElseIf cell = "学校" And InStr(nyTable(cell), "その他") > 0 Then
                isAllOvreWrite = isAllOvreWrite And True
            Else
                isAllOvreWrite = False
            End If
            
        End If
    Next
End Function

Private Function checkOption(ByRef argOptionCell As Object, ByVal argID As String) As Boolean
    Dim opt As Variant

    For Each opt In argOptionCell.Children
        If opt.id = argID Then
        
            If opt.Checked = False Then
                opt.click
                Exit For
            End If
        End If
    Next

End Function

Private Sub execNayose(argIWeb As CorpSite)
    Dim nyTable As Object 'Dictionary '
    Dim imptTable As Object 'Dictionary '
    Dim iwebTables As Collection
    Dim tb As Object
    
    Set imptTable = CreateObject("Scripting.Dictionary")
    Set iwebTables = New Collection
    
'##名寄せ画面から、テーブルを読み込み
    For Each tb In argIWeb.byTag("table")
        Set nyTable = loadNayoseTb(tb)
        
        If Not nyTable Is Nothing Then
            If InStr(tb.innerText, "更新(媒体のみ)") > 0 Then
                Set imptTable = nyTable
            Else
                iwebTables.Add nyTable
            End If
            
            Set nyTable = Nothing
        End If
    Next
    
    Dim chk As String
    
    chk = "chkbx_0"
    Set nyTable = iwebTables(1)
    
    opeLog.Add "★『" & imptTable(" 氏名") & "』を名寄せします。"
    
'##テーブルの有無確認
    If iwebTables.Count = 0 Then
        opeLog.Add "名寄候補がありません。"
        Exit Sub
    ElseIf iwebTables.Count > 1 Then
        opeLog.Add "複数の名寄せ候補があります！"
'        chk = "chkbx_" & iwebTables.Count - 1
'        Set nyTable = iwebTables(iwebTables.Count)
        Exit Sub
    End If

'##上書き対象と方法の選択
    Dim kubun As String
    
    If isAllOvreWrite(nyTable, imptTable) Then
        kubun = "nkbn2"
        opeLog.Add "上記より、ID : " & nyTable("ID") & " を全て更新します。"
    Else
        kubun = "nkbn4"
        opeLog.Add "上記より、ID : " & nyTable("ID") & " の媒体のみ更新します。"
    End If
    
    checkOption imptTable("選択"), kubun
    checkOption nyTable("選択"), chk '"chkbx_0"

'##上書き実施
    Dim btnVal As Variant
    
    For Each btnVal In Array("確　認", "実　行")
        For Each tb In argIWeb.byClass("s colored")
            If tb.Value = btnVal Then
                argIWeb.click tb
            End If
        Next
    Next
    
    Set nyTable = Nothing
    Set imptTable = Nothing
    Set iwebTables = Nothing
    
End Sub

Private Function sendMail(argIWeb As CorpSite) As Boolean
    Dim table As Variant
    Dim cntCells As Long
    Dim timeOut As Date
    Dim bufStr As String
    
    timeOut = SettingSh.DlTimeOut
    
    opeLog.Add "メールの予約を開始します。"
    
    With argIWeb
        If Not .submit(.byName("form1")(0), timeOut) Then GoTo err
        If Not .submit(.byName("formMake")(0), timeOut) Then GoTo err
        If Not .click(.byId("timeset"), timeOut) Then GoTo err
        
        For Each table In .byTag("table")
            If table.Cells(0).innerText = "予約可能時間" Then
                cntCells = table.Cells.Length
            
                If cntCells >= 13 Then
                    If Not .click(table.Cells(12).Children(0)) Then GoTo err
                Else
                    If Not .click(table.Cells(cntCells - 1).Children(0)) Then GoTo err
                End If
                
                Exit For
            End If
        Next
        
        'フォームに時刻が反映されるまで待機
        
        timeOut = Now + timeOut
        
        Do
            On Error Resume Next
            bufStr = .byId("timeno", , False).Value
            On Error GoTo 0
            
            If bufStr <> vbNullString Then
                Exit Do
            ElseIf Now > timeOut Then
                opeLog.Add "メール送信時刻設定でタイムアウトしました"
                GoTo err
            End If
            
            DoEvents
        Loop
        
        .byId("mailsflag-3").click
        
        If Not .click(.byId("btnConfirm"), timeOut) Then GoTo err
        If Not .click(.byId("btnComplete"), timeOut) Then GoTo err
    End With
    
    opeLog.Add "メールの予約が完了しました"
    sendMail = True
Exit Function

err:
    opeLog.Add "メールの予約に失敗しました"
    
End Function

Public Function ULPersonalData(argIWeb As CorpSite, ByVal filePath As String, ByVal iWebNo As Long) As Boolean
    Dim csvFiles As Collection
    Dim csvFilePath As Variant
    
    If Not loginiWeb(argIWeb) Then Exit Function
    
    filePath = verifyCSV(filePath)
    Set csvFiles = csvDivider(filePath, 500)
    
    If csvFiles.Count = 0 Then
        ULPersonalData = False
        Exit Function
    Else
        ULPersonalData = True
    End If
    
    For Each csvFilePath In csvFiles
        opeLog.Add "分割アップロード開始：" & csvFilePath
        ULPersonalData = ULPersonalData And ULPersonalOneFile(argIWeb, csvFilePath, iWebNo)
    Next
    
End Function

Private Function ULPersonalOneFile(argIWeb As CorpSite, ByVal filePath As String, ByVal iWebNo As Long) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    Dim i As Long
    Dim timeOut As Date
    Dim msg As String
    
    timeOut = SettingSh.DlTimeOut
    
    With argIWeb
        tgtURL = "wdi/index"
        pstData = ""
        
        If Not .post(.baseURL & tgtURL, , pstData) Then
            opeLog.Add "i-Web 一括インポート画面への遷移中に問題が発生しました。"
            Exit Function
        End If
    
        .byId("wdilayoutno").Value = iWebNo
        .byId("headerflg").click
        
        If Not setDialogByVBS Then Exit Function
        
        'ここの待ち時間が短いと次のファイル名をダイアログにインプットする動作でエラーがでる。
        DoEvents
        Sleep 500
    
        setFileName filePath
        
        Do While .byName("wdifile")(0).Value = vbNullString
             DoEvents
             If Now > timeOut Then Exit Do
        Loop
        
        Dim tgtElmt As Object
        Dim naviOk As Boolean
        
        '先頭データ確認ページへの遷移
        For Each tgtElmt In .byName("form_up")
            If InStr(tgtElmt.Action, "importfirstconfirm") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "先頭データのご確認をお願いします。") > 0 Then
                        naviOk = True
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " を取り込めません。" & vbCrLf & "先頭データ確認ページに遷移できません。"
            Exit Function
        Else
            naviOk = False
        End If

        '取込内容確認ページへの遷移
        For Each tgtElmt In .byName("form1")
            If InStr(tgtElmt.Action, "importsecondconfirm") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "処理内容のご確認をお願いします") > 0 Then
                        naviOk = True
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " を取り込めません。" & vbCrLf & "処理内容の確認ページに遷移できません"
            Exit Function
        Else
            naviOk = False
        End If
       
        '入力チェック結果のページへの遷移
        For Each tgtElmt In .byName("form1")
            If InStr(tgtElmt.Action, "filecheck") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "データチェック内容のご確認をお願いします") > 0 Then
                        naviOk = True
                    ElseIf InStr(.byClass("navigation_top")(0).innerText, "エラーデータが存在します") > 0 Then
                        opeLog.Add "★取り込みデータにエラーデータが存在します｡ 手動で取り込むか、エラーデータを除去してください。"
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " を取り込めません。" & vbCrLf & "入力チェック結果のページに遷移できません。"
            Exit Function
        Else
            naviOk = False
        End If
        
        '登録完了ページへの遷移
        For Each tgtElmt In .byName("form2")
            If InStr(tgtElmt.Action, "complete") > 0 Then
                If .submit(tgtElmt, timeOut) Then naviOk = True
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " を取り込めません。" & vbCrLf & "登録完了のページに遷移できません。"
            Exit Function
        Else
            naviOk = False
        End If
        
        Dim topMsg As String
        Dim formName As String
        
        '名寄せ/完了チェック
        Do While True
            Set tgtElmt = .byId("contents").getElementsByTagName("table")
            
            If tgtElmt.Length = 0 Then
                topMsg = .byId("contents").innerText
            Else
                topMsg = tgtElmt(0).innerText
            End If
            
            Set tgtElmt = Nothing
            
            If InStr(topMsg, "処理が完了しています") > 0 Then
                msg = " 新規登録：" & .getTableCell("新規登録件数", "新規登録件数", 1).innerText & "件 /" _
                     & " 更新件数：" & .getTableCell("更新件数", "更新件数", 1).innerText & "件 /" _
                     & " 無効データ件数：" & .getTableCell("無効データ件数", "無効データ件数", 1).innerText & "件"
            
                Exit Do
            ElseIf InStr(topMsg, "メール送信情報を登録いたしました") > 0 Then
                Exit Do
            
            ElseIf InStr(topMsg, "メール送信") > 0 Then
                msg = " 新規登録：" & .getTableCell("新規登録件数", "新規登録件数", 1).innerText & "件 /" _
                     & " 更新件数：" & .getTableCell("更新件数", "更新件数", 1).innerText & "件 /" _
                     & " 無効データ件数：" & .getTableCell("無効データ件数", "無効データ件数", 1).innerText & "件"
            
                If .MailFlg Then
                    sendMail argIWeb
                Else
                    Exit Do
                End If
                
            ElseIf InStr(topMsg, "削除") > 0 Then
                opeLog.Add "取り込みデータのエラーにより、インポートが完了しませんでした。"
                Exit Function
            
            ElseIf InStr(topMsg, "名寄せ") > 0 Then
                If InStr(topMsg, "更新いたしました。") > 0 Then
                    formName = "form"
                Else
                    formName = "form1"
                End If
    
                .submit .byName(formName)(0), timeOut
                
                execNayose argIWeb
            Else
                opeLog.Add "何らかの原因でインポートが完了しませんでした。" & vbCrLf _
                         & "i-Web画面のトップメッセージ:" & IIf(topMsg = vbNullString, "空白", topMsg)
                Exit Function
            End If
        Loop
    End With
        
    opeLog.Add msg
    mailInfo.Add msg
    
    ULPersonalOneFile = True
    
End Function

Public Function ULSeminarData(argIWeb As CorpSite, _
                                ByVal seminarJob As String, _
                                ByVal seminarName As String, _
                                ByVal uploadText As String, _
                                ByVal cancelFlg As Boolean) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    Dim classEl As Variant
    Dim param As Variant
    Dim msgText As String
    
    If uploadText = vbNullString Then
        'opeLog.Add "★" & seminarName & "に" & IIf(cancelFlg, "キャンセル", "登録") & "するIDはありません。"
        ULSeminarData = True
        Exit Function
    End If
    
    If Not loginiWeb(argIWeb) Then Exit Function
    If Not selectTab(argIWeb, seminarJob) Then
        opeLog.Add "★【失敗】i-Webトップ画面に『" & seminarJob & "』タブが見つかりません。" & vbCrLf & "以下の" & IIf(cancelFlg, "IDをキャンセル", "ID + [TAB] +イベントNo は登録") & "できておりません！！"
        opeLog.Add uploadText
        Exit Function
    End If
    
    With argIWeb
        For Each classEl In .byClass("clearfix")
            If classEl.innerText = seminarName Then
                param = Split(Split(Mid(classEl.outerHTML, InStr(classEl.outerHTML, "menu/event/") + 11), "'")(0), "/")
                pstData = param(0) & "=" & param(1) & "&" & param(2) & "=" & param(3)
            End If
        Next
        
        If pstData = vbNullString Then
            opeLog.Add "★【失敗】i-Webトップ画面の『" & seminarJob & "』タブ内に『" & seminarName & "』が見つかりません。" & vbCrLf & "以下の" & IIf(cancelFlg, "IDをキャンセル", "ID + [TAB] +イベントNo は登録") & "できておりません！！"
            opeLog.Add uploadText
            Exit Function
        End If
    
        tgtURL = "reserveimport/make"
        .post .baseURL & tgtURL, , pstData
        
        If cancelFlg Then
            .byId("eventdayno").Value = 99999
            .byId("matchingkey-1").click
            msgText = "★" & seminarJob & "/" & seminarName & "から以下のIDをキャンセルしました。" & vbCrLf
        Else
            .byId("matchingkey-2").click
            msgText = "★" & seminarJob & "/" & seminarName & "に以下のIDを予約しました。" & vbCrLf
        End If
        
        .byName("ucode").ucode.Value = uploadText
        .execScript "onSubmit();"
               
        If Trim(.getTableCell("エラー件数", "エラー件数", 2).innerText) <> "0件" Then
            opeLog.Add "★【失敗】" & seminarJob & "/" & seminarName & "へのアップロードに失敗しました。"
            opeLog.Add uploadText
            Exit Function
        End If
        
        .execScript "onSubmit();"

        'ログ出力整形
        Dim uploadArray As Variant
        Dim uploadDict As Object
        Dim v As Variant
        Dim pid As String
        Dim did As String
        Dim i As Long
        
        If Not cancelFlg Then
            Set uploadDict = CreateObject("Scripting.Dictionary")
            
            uploadArray = Split(uploadText, vbCrLf)
            uploadText = vbNullString
            
            For i = LBound(uploadArray) To UBound(uploadArray)
                pid = Split(uploadArray(i), vbTab)(0)
                did = Split(uploadArray(i), vbTab)(1)
                
                If Not uploadDict.Exists(did) Then
                    uploadDict.Add did, pid
                Else
                    uploadDict(did) = uploadDict(did) & vbCrLf & pid
                End If
            Next
            
            For Each v In uploadDict
                uploadText = uploadText & IIf(uploadText = vbNullString, vbNullString, vbCrLf) & "イベントNo:" & v & vbCrLf & uploadDict(v)
            Next
        End If
        
        opeLog.Add msgText & uploadText
        'mailInfo.Add msgText & uploadText
    End With
    
    ULSeminarData = True
    
End Function

Function ULAllSeminarData(ByRef iweb As CorpSite, ByRef myNavi As CorpSite, _
                          ByVal myNavSeminor As people, ByVal rkNavSeminar As people, ByVal iwebPeople As people) As Boolean
    Dim pid As Variant
    Dim semiList As SeminarList
    Dim tgtPerson As Person
    Dim hits As Collection
    Dim tgtPeople As Object 'Dictionary
    Dim tgtSeminar As Variant
    Dim hitsCnt As Long
    Dim altMailFlg As Boolean
    Dim ulFailFlg As Boolean
    Dim msg As String
    Dim topIdx As Long
    
    Set semiList = New SeminarList
    
    If Not semiList.setEvent(iweb.corpName) Then Exit Function
    
    Set tgtPeople = CreateObject("Scripting.Dictionary")
    
    mailInfo.Add vbCrLf & "＜セミナーインポート＞"
    
    topIdx = mailInfo.Count
    
    For Each tgtSeminar In Array(myNavSeminor, rkNavSeminar)
        If Not tgtSeminar Is Nothing Then
            For Each pid In tgtSeminar.allPeople
                Set tgtPerson = tgtSeminar.allPeople(pid)
                Set hits = iwebPeople.findPerson(tgtPerson)
                
                If hits.Count = 0 Then
                    Set hits = iwebPeople.findPerson(tgtPerson, True)
                    If hits.Count = 1 Then
                        With tgtPerson
                            msg = "★" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "は、i-Web上の" & hits(1).id & "とメールアドレスが不一致ですが、氏名・大学名・携帯電話番号が一致したため同一人物として扱います。"
                            opeLog.Add msg
                            mailInfo.Add msg
                            msg = vbNullString
                        End With
                    End If
                End If
                
                hitsCnt = hits.Count
                
                If hitsCnt = 1 Then
                    If Not tgtPeople.Exists(hits(1).id) Then
                        tgtPeople.Add hits(1).id, tgtPerson
                    Else
                        tgtPeople(hits(1).id).fusion tgtPerson
                    End If
                ElseIf hits.Count = 0 Then
                    With tgtPerson
                        opeLog.Add "★【失敗】" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "はi-Webに一致する人物が見つかりませんでした。"
                        opeLog.Add "メールアドレス：" & .mailAddress & " / " & .mobileAddress & vbCrLf & "電話番号：" & .mobileNumber & vbCrLf & "大学名：" & .university
                    End With
                    altMailFlg = True
                Else
                    With tgtPerson
                        opeLog.Add "★【失敗】" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "はi-Webに情報が一致する人物が複数おり、特定できません。" _
                                    & vbCrLf & "情報が一致する人物のID：" & hits(1).id & " " & "他" & hits.Count - 1 & "名"
                    End With
                    altMailFlg = True
                End If
            Next
        End If
    Next
    
    If altMailFlg Then msg = "ナビ側の個人情報とi - Web側の個人情報が一致せず､対象の人物が特定できない登録者がいます｡" & vbCrLf
    
    If tgtPeople.Count = 0 Then
        opeLog.Add "更新対象のセミナーログはありませんでした。"
        mailInfo.Add "対象者がいませんでした。"
    Else
        'iweb更新履歴を取り損ねるとFalseが返ってくる
        If Not semiList.setData(tgtPeople, iweb, myNavi) Then altMailFlg = True
        
        mailInfo.Add semiList.countMember & "名分のセミナー予約情報を更新しました。" & vbCrLf, after:=topIdx
        
        Dim jbName As Variant
        Dim smName As Variant
        Dim ulList As Variant
        
        Set ulList = semiList.outPutList
        
        For Each jbName In ulList
            For Each smName In ulList(jbName)
                '必ずキャンセルが先
                'キャンセルはステップ名単位で行われるので、同じステップ名の下で、イベントID①がキャンセル直後にイベントID②が予約、
                'となった場合、予約を先にしてしまうと、後のキャンセルの処理（イベントID無差別）でイベントID②がキャンセルされてしまう。
                '逆に、予約してすぐキャンセルした場合、同じイベントIDならfusion関数で上書きされるので問題無し、
                '違うIDならキャンセル後登録となるが、キャンセル対象は違うIDなので、該当イベントIDは予約状態（キャンセルされてない状態）で問題ないからOK。
                If Not ULSeminarData(iweb, jbName, smName, ulList(jbName)(smName)(bookState.Cancel), True) Then ulFailFlg = True
                If Not ULSeminarData(iweb, jbName, smName, ulList(jbName)(smName)(bookState.book), False) Then ulFailFlg = True
            Next
        Next
    End If
    
    If ulFailFlg Then
        msg = msg & "セミナー情報のアップロードに失敗した登録者がいます。" & vbCrLf
        altMailFlg = True
    End If
    
    On Error Resume Next
    loginiWeb iweb
    selectTab iweb, "全て表示"
    On Error GoTo 0
    
    If altMailFlg Then
        sendSemAlert msg & "当該の登録者はセミナーのアップロードを実行できておりません。" & vbCrLf & "詳細は実行ログを確認してください。"
    Else
        ULAllSeminarData = True
    End If

End Function

Public Function getIwebTopByXMLHTTP(ByVal baseURL As String, ByVal userName As String, ByVal userPass As String) As Object
    Dim objHTTP As Object
    Dim sendData As Variant
    Dim tgtAdd As String
    Dim resTxt As String

    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    With objHTTP
        sendData = "id=" & userName & "&pass=" & userPass
        tgtAdd = "login/check"
    
        .Open "POST", baseURL & tgtAdd, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send sendData
    End With

    If objHTTP.Status = 200 Then
        If InStr(objHTTP.responseText, "ログインできませんでした") Then GoTo loginErr
        Set getIwebTopByXMLHTTP = objHTTP
    Else
        GoTo loginErr
    End If
       
Exit Function
loginErr:
    opeLog.Add "★【失敗】通信エラーによりiWebにログインできませんでした。" & vbCrLf & "アドレス/アカウント/パスワード/通信状況をご確認ください。"

End Function

Public Function getIwebSeminarStateByXMLHTTP(ByVal baseURL As String, _
                                ByRef objHTTP As Object, _
                                ByVal tgtIwebId As String, _
                                ByVal tgtSeminarNo As String, _
                                ByVal lastUpdate As Date) As Long
    Dim sendData As Variant
    Dim tgtAdd As String
    Dim resTxt As String
    Dim tbl As Variant
    
    '最新の状況をデフォルトで「不明」※ナビ側の更新の方が新しい
    getIwebSeminarStateByXMLHTTP = bookState.Unknown

    With objHTTP
        sendData = "gkscd=" & tgtIwebId
        tgtAdd = "student/history/"
    
        .Open "POST", baseURL & tgtAdd, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send sendData
    End With

    If objHTTP.Status = 200 Then
        resTxt = objHTTP.responseText
    Else
        GoTo err
    End If
    
    'Dim hitFlg As Boolean
    Dim htmlDoc As Object
    
    Set htmlDoc = New HTMLDocument

    htmlDoc.write resTxt

    '更新履歴が無い場合は、終了。最新の状況は「不明」※ナビ側の更新の方が新しい
    If InStr(htmlDoc.getElementById("subContents").innerText, "更新履歴はありません") > 0 Then Exit Function
    
    'ログ表を取得
    Set tbl = htmlDoc.getElementById("selectTable")
    
    'ログ表が無い場合は想定外、エラー
    If tbl Is Nothing Then GoTo err
    
    '表から該当セミナーの最新のログ（不明、予約、キャンセル）を取得
    getIwebSeminarStateByXMLHTTP = getLastRecord(tbl, lastUpdate, tgtSeminarNo)

       
Exit Function
err:
    getIwebSeminarStateByXMLHTTP = -1
    opeLog.Add "★【失敗】iWebID : " & tgtIwebId & " は通信エラーによりi-Webのセミナーログが取得できませんでした。"

End Function

Private Function getLastRecord(ByVal logTable As Object, ByVal lastUpdate As Date, ByVal tgtSeminarNo As String) As Long
    Dim i As Long
    Dim n As Long
    Dim MultipleRows As Collection
    Dim logRow As Object
    Dim editBefore As String
    Dim editAfter As String
    Dim seminarJobStep As String
    
    seminarJobStep = SeminarSh.getSminarJobStep(tgtSeminarNo)
    
    Set MultipleRows = New Collection
       
    'ログを最後の行から追いかける（※昇順前提）
    For i = logTable.Rows.Length - 1 To 0 Step -1
        With logTable.Rows(i)
            '最初のセルが日付で、
            If IsDate(.Cells(0).innerText) Then
                If CDate(.Cells(0).innerText) < lastUpdate Then
                    'ナビ側最新日時より古ければ終了
                    '最新の状況は「不明」※ナビ側の更新の方が新しい
                    getLastRecord = bookState.Unknown
                    Exit Function
                Else
                    'ナビ側最新日時より新しい場合は保持
                    MultipleRows.Add logTable.Rows(i)
                End If
            Else
            '最初のセルが日付でないなら複数行のパターンなのでいったん保持して次の行へ移動
                MultipleRows.Add logTable.Rows(i)
                GoTo CONTINUE
            End If
        End With
        
        '最初のセルが日付で、ナビ側最新日時より新しい場合
        
        For Each logRow In MultipleRows
            With logRow
            'セル数が７の場合と４の場合で対象とするセルを変える
            n = .Cells.Length
            
            'ログテキストから、”i-WebイベントID”か”キャンセル”を抽出してくる。
                If n = 7 Then
                    editBefore = getEventIDfromLogText(.Cells(n - 4).innerText)
                    editAfter = getEventIDfromLogText(.Cells(n - 3).innerText)
                    
                ElseIf n = 4 Then
                    editBefore = getEventIDfromLogText(.Cells(n - 3).innerText)
                    editAfter = getEventIDfromLogText(.Cells(n - 2).innerText)
                End If
            End With
            
            '変更後に記載があるとき（ない時はスキップ）
            If editAfter <> vbNullString Then
                If editAfter = "キャンセル" Then
                    '変更後がキャンセルで、変更前の職種＋ステップが該当と等しい場合、最新の状況は「キャンセル」
                    '処理終了
                    If SeminarSh.getSminarJobStep(editBefore) = seminarJobStep Then
                         getLastRecord = bookState.Cancel
                         Exit Function
                    End If
                ElseIf SeminarSh.getSminarJobStep(editAfter) = seminarJobStep Then
                    '変更後がキャンセルでない、かつ変更前の職種＋ステップが該当と等しい場合、最新の状況は「予約」
                    '処理終了
                    getLastRecord = bookState.book
                    Exit Function
                End If
            End If
        Next
        
        '保持していた行をクリアにする。
        Set MultipleRows = New Collection
CONTINUE:
    Next
    
    '全ログに該当がなければ、最新の状況は「不明」※ナビ側の更新の方が新しい
    getLastRecord = bookState.Unknown
    
End Function

Private Function getEventIDfromLogText(ByVal LogText As String) As String
    Dim reg As Object
    Dim match As Object
    
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = "No\.?(\d+)[ :]|(キャンセル)"
    End With
    
    For Each match In reg.Execute(LogText)
        getEventIDfromLogText = IIf(IsEmpty(match.SubMatches(0)), match.SubMatches(1), match.SubMatches(0))
    Next
End Function
