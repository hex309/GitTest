VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Seminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sDayCode As String  'セミナー番号（マイナビ）/日程コード（リクナビ）
Private sStatus As Long   'キャンセルフラグ（マイナビ）/予約状態（リクナビ）
Private sUpdated As Date  'エントリー日（マイナビ）/個人予約情報更新日時（リクナビ）

Private corpName As String  '対象企業
Private seminarJobName  As String 'i-web 職種
Private seminarName  As String   'i-webイベント名称
Private iWebDayCode As String 'i-webイベントNo
Private myNavDayCode As String 'セミナー番号（マイナビ）
Private rkNavDayCode As String '日程コード（リクナビ）

'Private NoLogCancelFlg As Boolean
''マイナビで前回「予約」だったものが「キャンセル」になっており、かつエントリー日時が更新されてない場合True
    
Enum bookState
    book = 0
    Cancel = 1
    Unknown = 2
    UnloggedCancel = 99
End Enum

Public Property Get seminarDayCode() As String
    seminarDayCode = sDayCode
End Property

Public Property Get seminarStatus() As String
    seminarStatus = sStatus
End Property

Public Property Get seminarUpdated() As Date
    seminarUpdated = sUpdated
End Property

Public Property Get targetCorpName() As String
    targetCorpName = corpName
End Property

Public Property Get iWebSeminarJob() As String
    iWebSeminarJob = seminarJobName
End Property

Public Property Get iWebSeminarName() As String
    iWebSeminarName = seminarName
End Property

Public Property Get iWebEventNo() As String
    iWebEventNo = iWebDayCode
End Property

Public Property Get myNavEventNo() As String
    myNavEventNo = myNavDayCode
End Property

Public Property Get rkNavEventNo() As String
    rkNavEventNo = rkNavDayCode
End Property

'Public Property Get toCancelButNotUpdated() As Boolean
'    toCancelButNotUpdated = NoLogCancelFlg
'End Property

Public Property Let seminarDayCode(ByVal argDayCode As String)
    sDayCode = argDayCode
End Property

Public Property Let seminarStatus(ByVal argStatus As String)
    If IsNumeric(argStatus) Then
        If CLng(argStatus) <= 1 Then
            sStatus = argStatus
        ElseIf CLng(argStatus) = bookState.UnloggedCancel Then
            sStatus = bookState.UnloggedCancel
        Else
            sStatus = bookState.Unknown
        End If
    Else
        If argStatus = "予約" Then
            sStatus = bookState.book
        ElseIf argStatus = "キャンセル" Then
            sStatus = bookState.Cancel
        Else
            sStatus = bookState.Unknown
        End If
    End If
End Property

Public Property Let seminarUpdated(ByVal argUpdated As Date)
    sUpdated = argUpdated
End Property

Public Property Let targetCorpName(ByVal argCorpName As String)
    corpName = argCorpName
End Property

Public Property Let iWebSeminarJob(ByVal argSeminarJob As String)
    seminarJobName = argSeminarJob
End Property

Public Property Let iWebSeminarName(ByVal argSeminarName As String)
    seminarName = argSeminarName
End Property

Public Property Let iWebEventNo(ByVal argiWebEventNo As String)
    iWebDayCode = argiWebEventNo
End Property

Public Property Let myNavEventNo(ByVal argmyNavEventNo As String)
    myNavDayCode = argmyNavEventNo
End Property

Public Property Let rkNavEventNo(ByVal argrkNavEventNo As String)
    rkNavDayCode = argrkNavEventNo
End Property

'Public Property Let toCancelButNotUpdated(ByVal argIsCancel As Boolean)
'    NoLogCancelFlg = argIsCancel
'End Property

