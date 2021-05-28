Attribute VB_Name = "OpnRs_MOD"
Option Explicit
'★レコードセット取得系モジュール


Public Function Opn_ExlRs(ByVal str_StRng As String, str_Key As String, _
                                        Optional str_WHERE As String = "", _
                                        Optional str_Fild As String = "*", _
                                        Optional Flg As Long = 0)
'★Excelレコードセットオープン
    '(引数１:シート＆レンジ,引数２:主キー(Null除外フィールド)、引数３:追加条件文=省略時""
    '、引数４:フィールド指定=省略時"*",引数5:ヘッダー名有無し指定、１で無し＝省略時0で有）
    Set Exl_Cn = New ADODB.Connection
    Set Exl_Rs = New ADODB.Recordset
    Exl_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    If Flg = 0 Then
        Exl_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=YES;IMEX=1"
    ElseIf Flg = 1 Then
        Exl_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=NO;IMEX=1"
    End If
    Exl_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT " & str_Fild
    str_SQL = str_SQL & " FROM [" & str_StRng & "] " '管理表$A8:FZ8000
    str_SQL = str_SQL & " WHERE " & str_Key & " IS NOT NULL"
    Debug.Print str_SQL
    If str_WHERE <> "" Then
        str_SQL = str_SQL & str_WHERE
    End If
    Exl_Rs.Open str_SQL, Exl_Cn, adOpenKeyset, adLockReadOnly

End Function

Public Function Opn_AcRs(ByVal str_Tbl As String, str_Key As String, _
                                        Optional str_WHERE As String = "", _
                                        Optional str_Fild As String = "*", _
                                        Optional Flg As Long = 0)
'★Accessレコードセットオープン
'(引数１:テーブル名,引数２:Null除外フィールド名、引数３:追加条件文=省略時""、引数４:フィールド指定=省略時"*"）
    Set Ac_Cn = New ADODB.Connection
    Set Ac_Rs = New ADODB.Recordset
    str_AcDBcn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                          Sheets("ディレクトリ設定").Range("F8").Value & ";" 'AccessDBパスの取得と設定
'    On Error GoTo Era
    Ac_Cn.Open str_AcDBcn
    '◆Accessデータ抽出SQL
    str_SQL = ""
    str_SQL = str_SQL & " SELECT " & str_Fild
    str_SQL = str_SQL & " FROM " & str_Tbl
    If Flg = 0 Then
        str_SQL = str_SQL & " WHERE " & str_Key & " IS NOT NULL"
    End If
    If str_WHERE <> "" Then
        str_SQL = str_SQL & str_WHERE
    End If
    Debug.Print str_SQL
    Ac_Rs.Open str_SQL, Ac_Cn, adOpenForwardOnly, adLockPessimistic
    Exit Function
Era: 'エラー時処理*****************************************************
    If Err.Number = -2147467259 Then
        MsgBox "DBファイルへ接続できませんでした " & vbCrLf & _
         "ディレクトリ設定でパスを確認・再設定してください" & vbCrLf & _
         "OKを押すと設定ページへ移動します", 16
         Call vis_SETDirectSt
         End
    Else
        MsgBox Err.Number & vbCrLf & _
         Err.Description, 16
         End
    End If

End Function

Public Function Dis_Exl_Rs()
'★読出レコードセットのクローズと破棄
    On Error Resume Next
    If Exl_Rs Is Nothing Then
    Else
        Exl_Rs.Close
        Set Exl_Rs = Nothing
    End If
    If Exl_Cn Is Nothing Then
    Else
        Exl_Cn.Close
        Set Exl_Cn = Nothing
    End If
    
End Function

Public Function Dis_Ac_Rs()
'★読出レコードセットのクローズと破棄
    On Error Resume Next
    If Ac_Rs Is Nothing Then
    Else
        Ac_Rs.Close
        Set Exl_Rs = Nothing
    End If
    If Ac_Cn Is Nothing Then
    Else
        Ac_Cn.Close
        Set Ac_Cn = Nothing
    End If
    
End Function
