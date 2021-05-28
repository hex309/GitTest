VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "People"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mLabels As Object
Private mReqedLabels As Object
Private mPeople As Object
Private dLabels As Object


Private Sub Class_Initialize()
    Dim Label As Variant
    Dim cLabelDict As Object
    Dim ReqedLabels As Collection
    Dim i As Long, j As Long
    Dim srcSiteNames(1 To 3) As String
    
    Set mLabels = CreateObject("Scripting.Dictionary")
    Set mReqedLabels = CreateObject("Scripting.Dictionary")
    Set dLabels = CreateObject("Scripting.Dictionary")
    Set mPeople = CreateObject("Scripting.Dictionary")
   
    srcSiteNames(1) = "i-Web"
    srcSiteNames(2) = "マイナビ"
    srcSiteNames(3) = "リクナビ"
    
    Label = getCurrentRegion(LabelSh.Cells(1, 1), 1, False).Value
    
    For j = 1 To 3
        Set cLabelDict = CreateObject("Scripting.Dictionary")
        Set ReqedLabels = New Collection
    
        For i = LBound(Label, 1) To UBound(Label, 1)
            cLabelDict.Add Key:=Label(i, j), Item:=Label(i, 4)
            
            If Label(i, 6) Then
                ReqedLabels.Add Label(i, j)
            End If
        Next

        mLabels.Add Key:=srcSiteNames(j), Item:=cLabelDict
        mReqedLabels.Add Key:=srcSiteNames(j), Item:=ReqedLabels
        
        Set cLabelDict = Nothing
        Set ReqedLabels = Nothing
    Next
        
End Sub

Public Property Get allPeople() As Object
    Set allPeople = mPeople

End Property

Public Function setLabel(ByVal siteName As String, ByVal argLabel As Collection, Optional ByVal checkLabel As Boolean = True) As Boolean
    Dim reqedlabel As Variant
    Dim i As Long
    
    If checkLabel Then
        For Each reqedlabel In mReqedLabels(siteName)
            For i = 1 To argLabel.Count
                If reqedlabel = argLabel(i) Then
                    Exit For
                End If
            Next
            
            If i = argLabel.Count + 1 Then
                opeLog.Add "siteName : " & siteName & vbCrLf & "名寄せに必要な項目 『 " & reqedlabel & " 』がありません。"
                Exit Function
            End If
        Next
    End If

    If Not dLabels.Exists(siteName) Then
        dLabels.Add Key:=siteName, Item:=argLabel
        setLabel = True
    Else
        opeLog.Add "siteName : " & siteName & "のラベルは既に登録されています。"
    End If

    '重複の処理まだ
    'ラベル必須の存非チェックまだ
End Function

Public Function setData(ByVal siteName As String, ByVal argData As Collection, Optional ByVal dateFrom As Date = 0) As Boolean
    Dim i As Long
    Dim Person As Person
    
    Set Person = New Person
    
    If Not dLabels.Exists(siteName) Then
        opeLog.Add "siteName : " & siteName & " のラベルが読み込まれていません。"
        Exit Function
    End If

    For i = 1 To argData.Count
        Person.setData mLabels(siteName)(dLabels(siteName)(i)), argData(i)
    Next
    
    '更新日の1日前から読みこむ｡ (ナビのデータは1日1回朝に自動で全受信されるため､データが最大1日遅れている可能性がある)
    '秒以下の細かい影響を無視したいために、さらに1分さかのぼる
    If Person.Seminar.seminarUpdated >= DateAdd("n", -1, DateAdd("d", -1, dateFrom)) Then
        If Not mPeople.Exists(Person.id) Then
                mPeople.Add Key:=Person.id, Item:=Person
        Else
            If Not mPeople(Person.id).fusion(Person) Then
              opeLog.Add "ID : " & Person.id & " のセミナー情報の統合に失敗しました。"
              Exit Function
            End If
        End If
    End If

    
    setData = True
    
End Function

Public Function findPerson(ByVal tgtPerson As Person, Optional ByVal partialFlg As Boolean = False) As Collection
    Dim pid As Variant
    
    Set findPerson = New Collection
    
    For Each pid In mPeople
        If mPeople(pid).isTargetPerson2(tgtPerson) = IIf(partialFlg, 1, 2) Then
           findPerson.Add mPeople(pid)
        End If
    Next

End Function

Private Sub Class_Terminate()
    Set mLabels = Nothing
    Set dLabels = Nothing
    Set mPeople = Nothing
End Sub
