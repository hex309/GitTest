Attribute VB_Name = "Upd_MOD"
Option Explicit
'★SQLアップデート系モジュール

Public Sub upd_NewID()
'★新規ID登録更新
    Dim str_Cval As String
    
    str_Cval = Sheets("管理表新規登録").Range("D6").Value
    If CHK_Duplicate_ID("T_KANRI", str_Cval, "T_1") = True Then
        MsgBox "そのIDは既に使われています", 16, "重複エラー!"
        End
    End If
    Call Opn_AcRs("T_KANRI", "T_1")
    With Ac_Rs
        .AddNew
        !T_1 = str_Cval
        !RegDate = Now
        .Update
    End With
    Call Dis_Ac_Rs
  
End Sub

Public Sub Run_updKANRI()
'★管理表登録更新
    Call upd_KANRI
    
End Sub

Public Sub Run_updCostum_KANRI()

    Call upd_KANRI("管理表編集登録")
    
End Sub

Public Sub upd_KANRI(Optional ByVal str_Stn As String = "管理表編集登録", Optional Flg As Long = 0)
'★データ上書き更新処理
    '工程管理表の更新有無='有'のみ(T_KANRI)上書更新
    '外部データテーブル(T_GAIBU1,2)へは更新有無='有'のみ2キーユニーク(F_1&F_2)で上書
    Dim str_Fildn, str_Ans As String
    Dim i As Long
    Dim str_RngAd As String
    
    Call CHK_RegChange '更新有無判定
    str_Ans = Get_ChangeData '更新データID取得
    str_RngAd = Sheets(str_Stn).Range("B7").End(xlToRight).Address
    str_RngAd = Replace(str_RngAd, "7", "50")
    str_RngAd = Replace(str_RngAd, "$", "")
    Call Opn_ExlRs(str_Stn & "$B7:" & str_RngAd, "T_1", " AND RegFlg='有'")
    Call Opn_AcRs("T_KANRI", "T_1")
    With Ac_Rs '管理表上書更新開始
        Do Until .EOF
            If !T_1 = Exl_Rs!T_1 Then
                For i = 1 To Exl_Rs.Fields.Count - 1
                    str_Fildn = Exl_Rs.Fields(i).Name
                    Ac_Rs![RegFlg] = "更新有"
                    Ac_Rs![RegDate] = Now
                    Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
                    .Update
                Next i
            End If
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Call Dis_Ac_Rs
    Call upd_GAIB '外部データテーブルの上書更新
    Call Run_Search_KANRI 'データ再読
    MsgBox "データが更新されました !!" & vbCrLf & _
                  str_Ans, vbInformation
    Exit Sub
Era:
     
End Sub

Public Function upd_GAIB()
'★外部データの上書き更新
    Dim str_Fildn As String
    Dim str_SQLFild As String
    Dim i As Long
    
    Call Opn_ExlRs("管理表編集登録$B6:FZ8000", "F_1", " AND RegFlg='有'")
    Call Opn_AcRs("T_GAIBU1", "F_1")
    With Ac_Rs
        Do Until .EOF
            If !F_1 = Exl_Rs!F_1 Then
                If !F_2 = Exl_Rs!F_2 Then
                    For i = 0 To Exl_Rs.Fields.Count - 1
                        On Error GoTo Skip0
                        str_Fildn = Exl_Rs.Fields(i).Name
Skip0:
                        Ac_Rs![RegFlg] = "更新有"
                        If str_Fildn = "ID" Then GoTo Skip1  'シートにないカラムはスキップ
                        If str_Fildn = "ImpDate" Then GoTo Skip1 'シートにないカラムはスキップ
                        If str_Fildn = "RegFlg" Then GoTo Skip1 'シートにないカラムはスキップ
                        If InStr(str_Fildn, "_") <= 0 Then GoTo Skip1
                        Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
Skip1:
                    Next i
                    .Update
                End If
            End If
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Call Dis_Ac_Rs

End Function

Public Sub upd_KANRI_RegFlgRe()
'★管理表テーブルの更新有無を全てリセット
    'インポート毎に使用
    Dim sr_SQL As String
    
    Call Opn_AcRs("T_KANRI", "T_1")
    str_SQL = ""
    str_SQL = "UPDATE T_KANRI SET RegFlg=''"
    Set Ac_Cmd = New ADODB.Command
    With Ac_Cmd
        .ActiveConnection = Ac_Cn
        .CommandText = str_SQL
        .Execute
    End With
    Call Dis_Ac_Rs
    
End Sub
