Attribute VB_Name = "Sync_MOD"
Option Explicit
'★同期系モジュール
Dim str_AcDBcn As String

Public Sub Run_Douki_KANRIvew()
'★Ac⇒Ex管理表ビュー同期
    Call upd_AcExSync2("管理表出力ビュー", "T_KANRI", "B10", "T_1")

End Sub

Public Sub Run_Search_KANRI()
'★検索内容でAccessDBとExcelデータ管理表編集_登録シート同期
    Dim str_WHERE As String
  
    str_WHERE = Get_WHERE("管理表編集登録", "B2", "B4")
    Call upd_AcExSync2("管理表編集登録", "T_KANRI", "B10", "T_1", str_WHERE, 1)
    Call Re_Scrl

End Sub

Public Sub Run_Search_Costum_KANRI()
'★検索内容でAccessDBとExcelデータ管理表編集登録シート同期
    Call Run_Search_Costumvew("管理表編集登録")
    With Sheets("管理表編集登録")
        .Unprotect
        .Range("C:C").Formula = .Range("C:C").Formula
    End With
    Call St_Lock
    Call Re_Scrl
    ThisWorkbook.Save

End Sub

Public Sub Run_Search_KANRIvew(Optional ByVal str_Stn As String = "管理表出力ビュー", _
                                                     Optional str_rng As String)
'★検索内容でAccessDBとExcelデータ管理表出力ビューシート同期
    Dim str_WHERE As String
    
    Application.ScreenUpdating = False
    str_WHERE = Get_WHERE(str_Stn, "B2", "B4")
    Call upd_AcExSync2(str_Stn, "T_KANRI", "B10", "T_1", str_WHERE, 1)
    Call Re_Scrl

End Sub

Public Sub Run_Search_Costumvew(Optional ByVal str_Stn As String = "カスタムビュー")
'★検索内容でAccessDBとExcelデータ管理表出力ビューシート同期
    Dim str_WHERE As String
    Dim str_Fild As String
    Dim eCol As Long
    Dim i As Long
    
    eCol = ActiveSheet.Range("B7").End(xlToRight).Column
    str_Fild = ""
    For i = 2 To eCol
        str_Fild = str_Fild & ActiveSheet.Cells(7, i).Value & ","
    Next i
    str_Fild = Left(str_Fild, Len(str_Fild) - 1)
    str_WHERE = Get_WHERE(str_Stn, "B2", "B4")
    Call upd_AcExSync1(str_Stn, "T_KANRI", "T_1", "B10", str_Fild, str_WHERE)
    Call Re_Scrl

End Sub
Public Sub Run_Douki_GAIB()
'★Ac⇒Ex外部データ同期
    Dim str_Fild As String
    str_Fild = Get_SQLFelds("TG_G_ColList")
    Call upd_AcExSync1("外部データ", "T_GAIBU1", "F_1", "B8", str_Fild)

End Sub

Public Sub Run_Search_GAIB()
'★検索内容でAccessDBとExcelデータ管理表シート同期
    Dim str_WHERE As String
    Dim str_Fild As String
    str_Fild = Get_SQLFelds("TG_G_ColList")
  
    str_WHERE = Get_WHERE("外部データ", "B1", "B3")
    If str_WHERE = "" Then End
    
    Call upd_AcExSync1("外部データ", "T_GAIBU1", "F_1", "B8", str_Fild, str_WHERE)
    Call Re_Scrl

End Sub

Public Function upd_AcExSync1(ByVal str_Stn As String, str_Tbl As String, str_Key As String, _
                                                str_vRng As String, _
                                                Optional str_Fild As String = "*", _
                                                Optional str_WHERE As String = "")
'★Access外部データ⇒内部Excelシート同期
'　適用シート:カスタム
    '下記のupd_AcExSync2(管理表用)との違い:取得フィールド指定機能がある
    '(引数1:書き込みシート名,引数2:読出しテーブル名,引数3:Null除外フィールドセルアドレス
    '  引数4:貼付けセルアドレス,引数5: カラム設定シート名,引数6: 取得フィールド名=省略時は全フィールド,
    '引数7:SQL追加条件文=省略時は"")
    Dim L_Ws As Worksheet
    Dim str_LCStn As String
    Dim str_SQL As String
    Dim eRow As Long
    Dim RcCnt As String
'読出データセット Access
    Call Opn_AcRs("T_KANRI", str_Key, str_WHERE, str_Fild)
'読出データセットここまで
'◆データ転記
    Set L_Ws = Sheets(str_Stn)

    With L_Ws
        .Unprotect
         .Range("11:20000").Delete
        .Range("B11:GZ20000").ClearContents
        .Range(str_vRng).CopyFromRecordset Ac_Rs
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        If eRow < 10 Then End
        .Range("B10").EntireRow.Copy
        .Range(11 & ":" & eRow).PasteSpecial Paste:=xlPasteFormats
        .Range("G:HZ").EntireColumn.AutoFit
        RcCnt = eRow - 9
        If eRow = 10 Then
            .Range("11:11").Delete
            RcCnt = "1"
        End If
        If str_Stn = "管理表編集登録" Then
            .Range(eRow + 1 & ":1000").Interior.Color = 16777164
            .Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = RcCnt
        End If
    End With
    Call Dis_Ac_Rs
    Call St_Lock
    Exit Function
Era1:
    If Err.Number = -2147467259 Then
        MsgBox "DBファイルへ接続できませんでした " & vbCrLf & _
         "ディレクトリ設定でパスを確認・再設定してください" & vbCrLf & _
         "OKを押すと設定ページへ移動します", 16
         Call vis_SETDirectSt
        End
    End If
 
End Function

Public Function upd_AcExSync2(ByVal str_Stn As String, str_Tbl As String, _
                                                str_rng As String, str_Key As String, _
                                                Optional str_WHERE As String = "", Optional Flg As Long = 0)
'★Accessテーブルデータ⇒内部EXcelシート同期 (適用シート：管理表）
    '(引数1:書き込みシート名,引数2:読出しテーブル名,引数3:貼付けセルアドレス,引数4:Null判定列名
     ',引数5:検索キー文,全件フラグ =0:全件/=１:検索絞込 省略時は0)
    Dim L_Ws As Worksheet
    Dim str_SQL As String
    Dim sRow, sCol, eRow As Long
    '◆転記用シートオブジェクトセット
    Set L_Ws = Sheets(str_Stn)
    '◆検索時Access呼出データ有無チェック
    If Flg = 1 Then '全件フラグ=1=検索・絞込
        If str_WHERE <> "" Then '検索条件文指定あり
            Call Opn_AcRs(str_Tbl, str_Key, str_WHERE) '判定用Accessレコードセット
            If Ac_Rs.EOF = True Then 'データ有無判定＝データがなかったら
                MsgBox "データが見つかりませんでした", 16
                Call Dis_Ac_Rs
                End
            End If
        Else '検索条件文指定なし
            MsgBox "検索・絞込条件を指定してください", 16
            End
        End If
    ElseIf Flg = 0 Then '全件フラグ=0=全件
        Call Opn_AcRs(str_Tbl, str_Key)
    End If
    With L_Ws
        .Unprotect
        sRow = .Range(str_rng).Row
        .Range(sRow + 1 & ":90000").Delete
        .Range(str_rng & ":GZ10000").ClearContents
        .Range(str_rng).CopyFromRecordset Ac_Rs
        sCol = .Range(str_rng).Column
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        .Range("B10").EntireRow.Copy
        .Range(sRow + 1 & ":" & eRow).PasteSpecial Paste:=xlPasteFormats
        If eRow = 10 Then
            .Range("B11").EntireRow.Delete
        End If
        .Range(.Cells(sRow, sCol), .Cells(eRow, 10)).Formula = _
        .Range(.Cells(sRow, sCol), .Cells(eRow, 10)).Formula
        .Range(.Cells(Columns.Count, 7), .Cells(Columns.Count, 200)).EntireColumn.AutoFit
        .Range(eRow + 1 & ":1000").Interior.Color = 16777164
    End With
    Call Dis_Ac_Rs
    Call St_Lock

End Function
