VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeminarList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSeminar As Object
Private rkToIWeb As Object
Private myToIWeb As Object
Private jbNameList As Object
Private memberCnt As Long

Private Sub Class_Initialize()
    Set mSeminar = CreateObject("Scripting.Dictionary")
    Set rkToIWeb = CreateObject("Scripting.Dictionary")
    Set myToIWeb = CreateObject("Scripting.Dictionary")
    Set jbNameList = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set mSeminar = Nothing
    Set rkToIWeb = Nothing
    Set myToIWeb = Nothing
    Set jbNameList = Nothing
End Sub

Public Property Get outPutList() As Object
    Set outPutList = jbNameList
End Property

Public Property Get countMember() As Long
    countMember = memberCnt
End Property

Public Function setEvent(ByVal corpName) As Boolean
    Dim smlist As Variant
    Dim dCorpName As String
    Dim cSeminar As Seminar
    Dim i As Long

    On Error GoTo err
    smlist = getCurrentRegion(SeminarSh.Cells(2, 1), 1, False).Value
    On Error GoTo 0
    
    For i = LBound(smlist, 1) To UBound(smlist, 1)
        dCorpName = Trim(smlist(i, 1))
        
        If dCorpName = corpName Then
    
            Set cSeminar = New Seminar
        
            cSeminar.targetCorpName = dCorpName
            cSeminar.iWebSeminarJob = Trim(smlist(i, 2))
            cSeminar.iWebSeminarName = Trim(smlist(i, 3))
            cSeminar.iWebEventNo = Trim(smlist(i, 4))
            cSeminar.rkNavEventNo = Trim(smlist(i, 5))
            cSeminar.myNavEventNo = Trim(smlist(i, 6))
            
            If Not mSeminar.Exists(cSeminar.iWebEventNo) Then
                mSeminar.Add Key:=cSeminar.iWebEventNo, Item:=cSeminar
            Else
                opeLog.Add "日程コード : " & cSeminar.iWebEventNo & "　が重複しています。"
                GoTo err
            End If
            
            If Not rkToIWeb.Exists(cSeminar.rkNavEventNo) Then
                rkToIWeb.Add Key:=cSeminar.rkNavEventNo, Item:=cSeminar.iWebEventNo
            Else
                opeLog.Add "日程コード : " & cSeminar.rkNavEventNo & "　が重複しています。"
                GoTo err
            End If
            
            If Not myToIWeb.Exists(cSeminar.myNavEventNo) Then
                myToIWeb.Add Key:=cSeminar.myNavEventNo, Item:=cSeminar.iWebEventNo
            Else
                opeLog.Add "日程コード : " & cSeminar.myNavEventNo & "　が重複しています。"
                GoTo err
            End If
        End If
    Next
    
    setEvent = True
    
    Exit Function

err:
    opeLog.Add "対象セミナーリストが取得できませんでした。"

End Function

Public Function setData(ByVal people As Object, ByRef iweb As CorpSite, ByRef myNavi As CorpSite) As Boolean
    Dim pid As Variant
    Dim sid As Variant
    Dim iwebSid As String
    Dim tgtPerson As Person
    Dim smName As String
    Dim jbName As String
    Dim bookSt As Long
    Dim outLine As String
    Dim iwebX As Object 'MSXML2.XMLHTTP60
    Dim msg As String
    Dim smNameList As Object
    Dim smNoList As Object
    Dim notRegisteredEvents As String
    Dim naviSiteName As String
    Dim failFlag As Boolean
    Dim myNaviUpdate As Date
    Dim eventKey As String
    Dim eventKeys As Object
    
    'セミナーの重複を削る
    For Each pid In people
        Set tgtPerson = people(pid)
        
        Set eventKeys = CreateObject("Scripting.Dictionary")

        For Each sid In tgtPerson.Seminars
            iwebSid = getiWebEventNo(sid)
                       
            If iwebSid = vbNullString Then
                '登録外
                'do nothing
            Else
                '職種　＋　イベント名がキーになる
                eventKey = getEventJob(iwebSid) & getEventName(iwebSid)
                            
                '同じ職種＋イベント名がある場合
                If eventKeys.Exists(eventKey) Then
                    '更新日時が同じ場合は
                    If tgtPerson.Seminars(sid).seminarUpdated = tgtPerson.Seminars(eventKeys(eventKey)).seminarUpdated Then
                    
                        '新しいデータが「予約」の場合、セミナーリストから既存を削除し、キーの値を更新
                        If tgtPerson.Seminars(sid).seminarStatus = bookState.book Then
                            
                            With tgtPerson
                            opeLog.Add "【★要確認★】" & .id & .kanjiFamilyName & .kanjiFirstName & "のセミナー番号" & sid & "/" & eventKeys(eventKey) & _
                                       "の更新日時が" & Format(.Seminars(sid).seminarUpdated, "yyyy/mm/dd hh:mm") & " の同時刻で重複しています。" & _
                                       sid & "(i-WebイベントID：" & iwebSid & ")の「予約」を優先します。"
                            End With
                            
                            tgtPerson.Seminars.Remove eventKeys(eventKey)
                            eventKeys(eventKey) = sid
                            
                        '新しいデータが予約でない場合で、既存が予約なら新しいデータを破棄
                        ElseIf tgtPerson.Seminars(eventKeys(eventKey)).seminarStatus = bookState.book Then
                            
                            With tgtPerson
                            opeLog.Add "【★要確認★】" & .id & .kanjiFamilyName & .kanjiFirstName & "のセミナー番号" & sid & "/" & eventKeys(eventKey) & _
                                       "の更新日時が" & Format(.Seminars(sid).seminarUpdated, "yyyy/mm/dd hh:mm") & " の同時刻で複数あります。" & _
                                       eventKeys(eventKey) & "(i-WebイベントID：" & getiWebEventNo(eventKeys(eventKey)) & ")の「予約」を優先します。"
                            End With
                            
                            tgtPerson.Seminars.Remove sid
                        
                        '新しいデータ、既存のデータ共に「予約」ではない場、セミナーリストから既存を削除し、キーの値を更新
                        Else
                            tgtPerson.Seminars.Remove eventKeys(eventKey)
                            eventKeys(eventKey) = sid
                        End If
                    
                    '新しいデータの更新日時が新しい場合、セミナーリストから既存を削除し、キーの値を更新
                    ElseIf tgtPerson.Seminars(sid).seminarUpdated > tgtPerson.Seminars(eventKeys(eventKey)).seminarUpdated Then
                        tgtPerson.Seminars.Remove eventKeys(eventKey)
                        eventKeys(eventKey) = sid
                    
                    '新しいデータの更新日時が古い場合は新しいデータを破棄
                    Else
                        tgtPerson.Seminars.Remove sid
                    End If
                Else
                    eventKeys.Add eventKey, sid
                End If
            End If
        Next
    Next
    
    With iweb
        Set iwebX = getIwebTopByXMLHTTP(.baseURL, .userName, .userPass)
    End With
       
    'ログインできなかった場合は異常終了
    If iwebX Is Nothing Then Exit Function
        
    For Each pid In people
        Set tgtPerson = people(pid)
        
        For Each sid In tgtPerson.Seminars
            iwebSid = getiWebEventNo(sid)
                        
            If iwebSid = vbNullString Then
                bookSt = tgtPerson.Seminars(sid).seminarStatus
                
                If bookSt = bookState.book Then
                    msg = "予約"
                ElseIf bookSt = bookState.Cancel Then
                    msg = "キャンセル"
                Else
                    msg = "不明"
                End If
                
                bookSt = 0
                
                Select Case Len(sid)
                    Case 5
                        naviSiteName = "リクナビ"
                    Case 15
                        naviSiteName = "マイナビ"
                    Case Else
                        naviSiteName = "ナビサイト不明"
                End Select
                
                notRegisteredEvents = notRegisteredEvents & IIf(notRegisteredEvents = vbNullString, vbNullString, vbCrLf) _
                                        & naviSiteName & "/" & sid & "/" & tgtPerson.id & "/" & pid & "/" & msg
                msg = vbNullString
                
                GoTo CONTINUE
            End If
            
            smName = getEventName(iwebSid)
            jbName = getEventJob(iwebSid)
            
            bookSt = getIwebSeminarStateByXMLHTTP(iweb.baseURL, iwebX, pid, iwebSid, tgtPerson.Seminars(sid).seminarUpdated)
            'iwebの更新履歴が取得できなかった場合は、フラグ立ててスキップ
            If bookSt = -1 Then
                failFlag = True
                GoTo CONTINUE
            End If
            
            'ログが更新されていないキャンセルに対する処理
            If tgtPerson.Seminars(sid).seminarStatus = bookState.UnloggedCancel Then
                If bookSt = bookState.Cancel Or bookSt = bookState.Unknown Then
                    tgtPerson.Seminars(sid).seminarStatus = bookState.Cancel
                    opeLog.Add "【★マイナビログ確認済★】" & tgtPerson.id & "/" & pid & ":" & tgtPerson.kanjiFamilyName & " " & tgtPerson.kanjiFirstName & " の" & smName & _
                               "は、前回更新日時から現在までのどこかで、マイナビ上でキャンセルされています。 " & vbCrLf & _
                               "以下のいずれかの条件に当てはまるため最新の状態をキャンセルとします。" & vbCrLf & _
                               "・i-Web側の該当セミナーの履歴が、ツールの前回更新日時の一日前より古い。" & vbCrLf & _
                               "・i-Web側の該当セミナーの履歴が、リクナビの予約日時より古い。" & vbCrLf & _
                               "・i-Web側の該当セミナーのステータスがキャンセル。"
                Else
                    'i-webが前回以降のログで「予約」、ナビ側が前回「予約」⇒今回「キャンセル」されているケース（エントリ―日時は前回以前のままで）
                    'つまり、「予約」と「キャンセル」のいずれが最新かわからないケース。
                    msg = "【★マイナビログ確認済★】" & tgtPerson.id & "/" & pid & ":" & tgtPerson.kanjiFamilyName & " " & tgtPerson.kanjiFirstName & " の" & smName & _
                          "は、前回更新日時から現在までのどこかで、マイナビ上でキャンセルされています。 " & vbCrLf & _
                          "またマイナビとiWebの状態が一致しておらず、かつどちらが最新か判別できないためマイナビのキャンセル日時を確認します。"
                    opeLog.Add msg

                    myNavi.visible True
                    myNaviUpdate = getMynaviCancelDate(myNavi, tgtPerson.id, sid, 2)
                    
                    If myNaviUpdate = 0 Then
                        opeLog.Add "マイナビ側にログがないか、最新のステータスが「予約」です。i-Web側も「予約」であるためスキップします。"
                        GoTo CONTINUE
                    Else
                        opeLog.Add "マイナビの更新日時は" & myNaviUpdate & "でした。"
                        opeLog.Add "i-web側と比較しナビ側が新しい場合はi-webをキャンセルします。"
                    End If
                    
                    tgtPerson.Seminars(sid).seminarStatus = bookState.Cancel
                    bookSt = getIwebSeminarStateByXMLHTTP(iweb.baseURL, iwebX, pid, iwebSid, myNaviUpdate)

                End If
            End If
            
            If bookSt = tgtPerson.Seminars(sid).seminarStatus Then
                If bookSt = bookState.book Then
                    msg = "既に『予約』済み"
                ElseIf bookSt = bookState.Cancel Then
                    msg = "既に『キャンセル』済み"
                Else
                    msg = "不明"
                End If
                
                msg = "★" & tgtPerson.id & "/" & pid & ":" & tgtPerson.kanjiFamilyName & " " & tgtPerson.kanjiFirstName & " の" & smName & "は、" & msg & "の為スキップします。"
                opeLog.Add msg
                'mailInfo.Add msg
                
            ElseIf bookSt <> bookState.Unknown Then
                msg = "★" & tgtPerson.id & "/" & pid & ":" & tgtPerson.kanjiFamilyName & " " & tgtPerson.kanjiFirstName & " の" & smName & "は、i-web側の更新日時が新しい為スキップします。"
                opeLog.Add msg
                'mailInfo.Add msg
            Else
                bookSt = tgtPerson.Seminars(sid).seminarStatus
                
                If jbNameList.Exists(jbName) Then
                    Set smNameList = jbNameList(jbName)
                Else
                    Set smNameList = CreateObject("Scripting.Dictionary")
                    jbNameList.Add jbName, smNameList
                End If
                         
                If smNameList.Exists(smName) Then
                    Set smNoList = smNameList(smName)
                Else
                    Set smNoList = CreateObject("Scripting.Dictionary")
                    smNoList.Add bookState.book, vbNullString
                    smNoList.Add bookState.Cancel, vbNullString

                    smNameList.Add smName, smNoList
                End If
                
                If bookSt = bookState.Cancel Then
                    smNoList(bookSt) = IIf(smNoList(bookSt) = vbNullString, vbNullString, smNoList(bookSt) & vbCrLf) & pid
                Else
                    smNoList(bookSt) = IIf(smNoList(bookSt) = vbNullString, vbNullString, smNoList(bookSt) & vbCrLf) & pid & vbTab & iwebSid
                End If
                    
                Set smNameList(smName) = smNoList
                Set jbNameList(jbName) = smNameList
                memberCnt = memberCnt + 1
            End If
CONTINUE:
        Next
    Next
    
    
    If Not notRegisteredEvents = vbNullString Then
       notRegisteredEvents = "★下記はイベントの更新がありましたが、セミナー対照表に記載がない為、i-webへの更新をしていません。" & vbCrLf _
                                & "ナビサイト名/セミナID/ナビID/iWebID/状況" & vbCrLf _
                                & notRegisteredEvents
        mailInfo.Add notRegisteredEvents
        opeLog.Add notRegisteredEvents
    End If
    
    setData = Not failFlag
    
End Function

Public Function getEventName(ByVal argiWebEventNo As String) As String
    On Error Resume Next
    getEventName = mSeminar(argiWebEventNo).iWebSeminarName
    On Error GoTo 0
End Function

Public Function getEventJob(ByVal argiWebEventNo As String) As String
    On Error Resume Next
    getEventJob = mSeminar(argiWebEventNo).iWebSeminarJob
    On Error GoTo 0
End Function

Private Function getiWebEventNo(ByVal argEventNo As String) As String
    
    If myToIWeb.Exists(argEventNo) Then
            getiWebEventNo = myToIWeb(argEventNo)
    ElseIf rkToIWeb.Exists(argEventNo) Then
            getiWebEventNo = rkToIWeb(argEventNo)
    Else
        Exit Function
    End If
    
End Function
