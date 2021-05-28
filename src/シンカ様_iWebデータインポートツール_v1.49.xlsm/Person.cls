VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================================
'Project Name    :
'File Name       : .xlsm
'Creation Date   : 2018/05/29
'
'Copyright (c) 2018 CriticMission Japan Co., Ltd. All rights reserved.
'======================================================================

'// ************************************************************************************************************************
'//  モジュール名　　：Person
'//  概要　　　　　　：
'//  参照設定　　　　：なし
'//  作成者　　　　　：Akira Hashimoto
'//  作成日　　　　　：2018/05/29
'//  更新日：内容　　：
'// ************************************************************************************************************************

    Private apCode As String    '応募者コード
    Private kjFamName As String '漢字姓
    Private kjFstName As String '漢字名
    Private knFamName As String 'カナ姓
    Private knFstName As String 'カナ名
    Private sex As String       '性別
    Private birthDate As Date   '生年月日
    Private univName As String  '大学名称
    Private deptName As String  '学部名称
    Private coursName As String '学科名称
    Private postCode As String  '現住所郵便番号
    Private addPref As String   '現在住所都道府県
    Private addCity As String   '現在住所市区郡番地
    Private addCho As String    'なし（現住所町名）
    Private addBanch As String  'なし（現住所番地）
    Private addBuild As String  '現住所マンション名
    Private vcTelNum As String  '休暇中TEL
    Private moTelNum As String  '携帯電話番号
    Private mail As String      '電子メールアドレス
    Private moMail As String    '携帯メールアドレス
    
    Private dSeminar As Seminar 'セミナー日程
    Private Schedule As Object 'セミナー日程のディクショナリ key:日程コード　item:seminar

Private Sub Class_Initialize()
    Set Schedule = CreateObject("Scripting.Dictionary")
    Set dSeminar = New Seminar

End Sub

Private Sub Class_Terminate()
    Set Schedule = Nothing
    Set dSeminar = Nothing
    
End Sub

Public Property Get id() As String
    id = apCode
End Property

Public Property Get kanjiFamilyName() As String
    kanjiFamilyName = kjFamName
End Property

Public Property Get kanjiFirstName() As String
    kanjiFirstName = kjFstName
End Property

Public Property Get kanaFamilyName() As String
    kanaFamilyName = knFamName
End Property

Public Property Get kanaFirstName() As String
    kanaFirstName = knFstName
End Property

Public Property Get mailAddress() As String
    mailAddress = mail
End Property

Public Property Get mobileAddress() As String
    mobileAddress = moMail
End Property

Public Property Get vcTelNumber() As String
    vcTelNumber = vcTelNum
End Property

Public Property Get mobileNumber() As String
    mobileNumber = moTelNum
End Property

Public Property Get university() As String
    university = univName
End Property

Public Property Get univGakubu() As String
    univGakubu = deptName
End Property

Public Property Get univGakka() As String
    univGakka = coursName
End Property

Public Property Get Seminar() As Seminar
    Set Seminar = dSeminar
End Property

Public Property Get Seminars() As Object
    Set Seminars = Schedule
End Property


Public Function setData(ByVal Label As String, ByVal newData As String) As Boolean
       
    'データ抜けチェック
    If Label = vbNullString Or newData = vbNullString Then
        'データが無い場合は初期値のまま抜ける。各項目の型により初期値は異なる。
        setData = True
        Exit Function
    End If
    
    On Error GoTo setDataErr:
    'データをプロパティにセット。型変換できない場合はエラーで抜ける。
    
    Select Case Label
        Case "AP_CODE": apCode = newData
        Case "KJ_FAM_NAME": kjFamName = newData
        Case "KJ_FST_NAME": kjFstName = newData
        Case "KN_FAM_NAME": knFamName = newData
        Case "KN_FST_NAME": knFstName = newData
        Case "SEX": sex = newData
        Case "BIRTH_DATE": birthDate = CDate(newData)
        Case "UNIV_NAME": univName = newData
        Case "DEPT_NAME": deptName = newData
        Case "COURS_NAME": coursName = newData
        Case "POST_CODE": postCode = newData
        Case "ADD_PREF": addPref = newData
        Case "ADD_CITY": addCity = newData
        Case "ADD_CHO": addCho = newData
        Case "ADD_BANCHI": addBanch = newData
        Case "ADD_BUILD": addBuild = newData
        Case "VC_TEL_NUM": vcTelNum = newData
        Case "MO_TEL_NUM": moTelNum = newData
        Case "MAIL": mail = newData
        Case "MO_MAIL": moMail = newData
        Case "S_DAY__CODE"
                dSeminar.seminarDayCode = newData
                Schedule.Add Key:=newData, Item:=dSeminar
        Case "S_STATUS": dSeminar.seminarStatus = newData
        Case "S_UPDATED": dSeminar.seminarUpdated = CDate(newData)
        Case Else
            opeLog.Add "ラベル：" & Label & " は規定されていません。" & vbCrLf & "ラベルに誤りがないか確認してください。"
            setData = False
            Exit Function
    End Select
    
    setData = True
    
    On Error GoTo 0
Exit Function

setDataErr:
    If err.Number = 13 Then
        opeLog.Add "『　" & newData & " 』" & vbCrLf _
                & "上記の値は、ラベル：" & Label & " の値として変換できません。" & vbCrLf _
                & "値を確認してください。"
    Else
        opeLog.Add "『　" & newData & " 』" & vbCrLf _
                & "ラベル：" & Label & " " & vbCrLf _
                & err.Description
    End If
    
    setData = False
    
End Function

Public Function fusion(ByVal newMySelf As Person) As Boolean
    Dim sid As Variant
    
    '新しいものを優先します。
    
    '同じIDの新しい行（新しい自分）が来た場合
    '今の自分のスケジュールに同じ日程IDのセミナーがあるか確認
    
    For Each sid In newMySelf.Seminars
        If Schedule.Exists(sid) Then
            'ある場合は、今のセミナーの最終更新日時と、新しい自分が持っている同じセミナーの更新日時を比較
            
            '新しい自分の更新日時が同じ場合、「予約」を優先
            If newMySelf.Seminars(sid).seminarUpdated = Schedule(sid).seminarUpdated Then

                If newMySelf.Seminars(sid).seminarStatus = bookState.book Or Schedule(sid).seminarStatus = bookState.book Then
                    opeLog.Add "【★要確認★】" & Me.id & Me.kanjiFamilyName & Me.kanjiFirstName & "のセミナー番号" & sid & _
                               "の更新日時が" & Format(Schedule(sid).seminarUpdated, "yyyy/mm/dd hh:mm") & " の同時刻で複数あります。" & _
                               "「予約」を優先します。"
                    Schedule(sid).seminarStatus = bookState.book
                Else
                    '「予約」が無い場合は、後を優先｡
                    Set Schedule(sid) = newMySelf.Seminar
                End If
                
            ElseIf newMySelf.Seminars(sid).seminarUpdated > Schedule(sid).seminarUpdated Then
                '新しい自分の更新日時が新しい場合、セミナーを上書き。
                Set Schedule(sid) = newMySelf.Seminar
            Else
                'Do Nothing
            
            End If
        Else
            '同じ日程IDがなければ、スケジュールに追加
            Schedule.Add Key:=sid, Item:=newMySelf.Seminars(sid)
        End If
    Next
    
    fusion = True

End Function

Public Function isTargetPerson2(ByVal tgtPerson As Person) As Long
    '##氏名、メールアドレス、電話番号を確認。
    '##メールアドレス2つのうちいずれか一致で 2, メール不一致かつ、他一致で 1, 他は0 を返す
    
    With tgtPerson
        If isTargetPerson(, , , , , , , .mailAddress, .mobileAddress, True) Then
            isTargetPerson2 = 2
        ElseIf isTargetPerson(.kanjiFamilyName, .kanjiFirstName, .university, , , , .mobileNumber) Then
            isTargetPerson2 = 1
        Else
            isTargetPerson2 = 0
        End If
    End With
End Function

Public Function isTargetPerson(Optional ByVal argKjFamName As String, _
                                Optional ByVal argKjFstName As String, _
                                Optional ByVal argUnivName As String, _
                                Optional ByVal argDeptName As String, _
                                Optional ByVal argCoursName As String, _
                                Optional ByVal argVcTelNum As String, _
                                Optional ByVal argMoTelNum As String, _
                                Optional ByVal argMail As String, _
                                Optional ByVal argMoMail As String, _
                                Optional ByVal orFlg As Boolean = False) As Boolean
                                
    Dim allBlankFlg As Boolean: allBlankFlg = True
    
    '氏＋名が一致で通す
    If argKjFamName <> vbNullString And argKjFstName <> vbNullString Then
        If argKjFamName & argKjFstName = kjFamName & kjFstName Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If argUnivName <> vbNullString Then
        If argUnivName = univName Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If argDeptName <> vbNullString Then
        If argDeptName = deptName Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If argCoursName <> vbNullString Then
        If argCoursName = coursName Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If argVcTelNum <> vbNullString Then
        If argVcTelNum = vcTelNum Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If argMoTelNum <> vbNullString Then
        If argMoTelNum = moTelNum Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    'メールはモバイルとPCの区別があいまいなため、いずれか一致で通す
    If argMail <> vbNullString Then
        If argMail = mail Or argMail = moMail Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    'メールはモバイルとPCの区別があいまいなため、いずれか一致で通す
    If argMoMail <> vbNullString Then
        If argMoMail = moMail Or argMoMail = mail Then
            allBlankFlg = False
        Else
            If Not orFlg Then Exit Function
        End If
    End If
    
    If allBlankFlg Then
        Exit Function
    Else
        isTargetPerson = True
    End If
    
End Function
