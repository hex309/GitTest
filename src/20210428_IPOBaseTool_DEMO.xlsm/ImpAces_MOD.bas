Attribute VB_Name = "ImpAces_MOD"
Option Explicit
'★Accessインポート系モジュール

Public Sub ins_Ex_Ac()
'★Excel⇒AccessDB外部データインサート 200x2
    Call ins_St_Tbl("TMP_R1$A1:GR", "T_GAIBU1")
    Call ins_St_Tbl("TMP_R2$A1:GR", "T_GAIBU2")

End Sub

Public Sub ins_Ex_Ac_Kanri_CSV()
'★Excel⇒AccessDB管理表バックアップCSVデータインサート 202
    Call ins_St_Tbl("TMP_CSV$A1:GT", "T_KANRI", "TMP_CSV", 3, "T_1")

End Sub

Public Function ins_St_Tbl(ByVal str_rng As String, str_Tbl As String, _
                                        Optional str_Stn As String = "一括取込", _
                                        Optional RowCnt As Long = 1, _
                                        Optional KeyCol As String = "F_1", Optional Flg As Long = 0)
'★Excelデータ⇒Accessインサート
    '→指定シートから指定テーブルへデータをインサート(引数１シート名と範囲、引数２テーブル名
    ',テーブルデリートフラグ 0=毎回デリート　1=デリート無し　デフォルト=0)
    Dim i, eRow As Long
    Dim str_Fildn As String
    Dim WHword As String
    
    eRow = Sheets(str_Stn).Cells(Rows.Count, RowCnt).End(xlUp).Row '一括取込シート最終行取得
    Call Opn_ExlRs(str_rng & eRow, KeyCol) '読出データセット Excel
    Call Opn_AcRs(str_Tbl, KeyCol) '書込データセット Access
    Debug.Print Ac_Rs.State
'    On Error GoTo Era
    If Flg = 0 Then 'Flg=０でテーブル毎回クリア
        Set Ac_Cmd = New ADODB.Command
        str_SQL = ""
        str_SQL = str_SQL & "DELETE FROM " & str_Tbl
        With Ac_Cmd
            .ActiveConnection = Ac_Cn
            .CommandText = str_SQL
            .Execute
        End With
    End If
    Debug.Print Ac_Rs.State

    With Ac_Rs 'データ上書開始
        Do Until Exl_Rs.EOF
            .AddNew
            For i = 0 To Exl_Rs.Fields.Count - 1
                 str_Fildn = Exl_Rs.Fields(i).Name
                 If str_Fildn = "ImpDate" Then
                    ![ImpDate] = Now
                ElseIf str_Fildn = "RegDate" Then
                    ![RegDate] = Now
                End If
                Ac_Rs(str_Fildn).Value = Exl_Rs(str_Fildn).Value
            Next i
            .Update
            Exl_Rs.MoveNext
        Loop
    End With
    Set Ac_Cmd = New ADODB.Command
    If KeyCol = "F_1" Then
        WHword = "あ" '
    ElseIf KeyCol = "T_1" Then
        WHword = "管理表ID"
    End If
    str_SQL = ""
    str_SQL = "DELETE FROM " & str_Tbl & " WHERE " & KeyCol & "='" & WHword & "'"  'データ型自動変換対策用レコードの削除
    With Ac_Cmd
        .ActiveConnection = Ac_Cn
        .CommandText = str_SQL
        .Execute
    End With
    Call Dis_Ac_Rs
    Call Dis_Exl_Rs
    Exit Function
'エラー時処理 ******************************************
Era:
     If Err.Number = -2147467259 Then
        MsgBox "DBファイルへ接続できませんでした " & vbCrLf & _
         "ディレクトリ設定でパスを確認・再設定してください" & vbCrLf & _
         "OKを押すと設定ページへ移動します", 16
        Call Dis_Ac_Rs
        Call Dis_Exl_Rs
        Call vis_SETDirectSt
        End
    Else
        MsgBox "エラー" & vbCrLf & _
        Err.Description, 16
        Call Dis_Ac_Rs
        Call Dis_Exl_Rs
        End
    End If

End Function
