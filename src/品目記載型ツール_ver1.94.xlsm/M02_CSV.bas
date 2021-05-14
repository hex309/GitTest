Attribute VB_Name = "M02_CSV"
Option Explicit

Public Pub対象ファイルパス As String
Public Pub件名 As String
Public Pub見積番号 As String
Public Pub建業法 As String
Public PubHas建業法 As Boolean
Public Pub下請法 As String
Public PubHas下請法 As Boolean

Public Pub日付 As Variant
Public Pub社員番号  As Variant

'CSVファイル作成
Public Sub CSVファイル作成(ByVal Target As Range)
    Dim タイトル As Variant
    Dim 種別フラグ As String
    Dim 見積もり件名 As String
    Dim パターン As String
    Dim 開始日 As Variant
    Dim 終了日 As Variant
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHdl
    
    タイトル = Target.Offset(, 2).Value
    種別フラグ = Target.Offset(, 3).Value
    Dim 回線合算フラグ As Boolean
    If Target.Offset(0, 3).Value Like "*回線合算*" Then
        回線合算フラグ = True
    Else
        回線合算フラグ = False
    End If
    パターン = Target.Offset(, 5).Value
    Dim ビジネスIT As String
    ビジネスIT = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE).Range("AU3").Value
    '    見積もり件名 = Target.Offset(, 4).Value
    
    If 必須データチェック(Target, パターン) = False Then
        Exit Sub
    End If
    
    
    If IS_TEST = False Then
        制御文解析 WSNAME_LOGIN
    End If
    '
    '保存先チェック
    Dim SavePath As String
    SavePath = 設定取得("▼品目記入型", "csvファイル保存先フォルダ")
    If HasTargetFolder(SavePath) = False Then
        MsgBox "csvファイル保存先フォルダを「設定」シートに指定してください" _
            , vbExclamation
        Exit Sub
    End If

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    sh.Rows(3).Clear
    Dim vRow As Long
    vRow = sh.Cells.SpecialCells(xlCellTypeLastCell).Row
    If vRow < 6 Then
        vRow = 6
    End If
    sh.Rows("6:" & sh.Rows.Count).Clear
    '    sh.Range("A5").CurrentRegion.Offset(1).Clear
    Dim 見積前提条件 As String
    ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value = ""

    モジュール転記 パターン

    Dim i As Long, j As Long
    i = 6
    Do
        DoEvents
        If sh.Cells(i, 1).Value = "" Then
            Exit Do
        End If
        品目情報転記 sh.Cells(i, 1).Value, i, 回線合算フラグ
        i = i + 1
    Loop

    Dim 変換対象 As Range
    Dim LastRow As Long
    With Sheet0
        LastRow = .Cells(.Rows.Count, 15).End(xlUp).Row
        If LastRow < 6 Then
            LastRow = 6
        End If
        Set 変換対象 = .Range(.Cells(6, 15), .Cells(LastRow, 19))
    End With
    
    Dim HasKengyoHo As Boolean
    Dim HasShitaukeHo As Boolean
    Dim EndRow As Long
    Dim vBook As Workbook
    '    Set vBook = Workbooks.Add
    '    ThisWorkbook.Worksheets(WSNAME_FORMAT).Cells.Copy _
    '        vBook.Worksheets(1).Range("A1")
    Dim 見積シート As Worksheet
    Dim vMsg As String
    Set 見積シート = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    For i = 1 To 変換対象.Rows.Count
        DoEvents
        HasKengyoHo = False
        Pub建業法 = vbNullString
        Pub下請法 = vbNullString
        見積もり件名 = Target.Offset(, 4).Value
        見積前提条件 = ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value

        If Target.Offset(, 6).Value = "" Then
            見積もり件名 = Replace(見積もり件名, "★★", 変換対象(i, 1))
            見積前提条件 = Replace(見積前提条件, "★★", 変換対象(i, 1))
        Else
            見積もり件名 = Replace(見積もり件名, "★★", Target.Offset(, 6).Value)
            見積前提条件 = Replace(見積前提条件, "★★", Target.Offset(, 6).Value)
        End If
        If Target.Offset(, 7).Value = "" Then
            見積もり件名 = Replace(見積もり件名, "●●", 変換対象(i, 2))
            見積前提条件 = Replace(見積前提条件, "●●", 変換対象(i, 2))
        Else
            見積もり件名 = Replace(見積もり件名, "●●", Target.Offset(, 7).Value)
            見積前提条件 = Replace(見積前提条件, "●●", Target.Offset(, 7).Value)
        End If
        If Target.Offset(, 8).Value = "" Then
            見積もり件名 = Replace(見積もり件名, "■■", 変換対象(i, 3))
            見積前提条件 = Replace(見積前提条件, "■■", 変換対象(i, 3))
        Else
            見積もり件名 = Replace(見積もり件名, "■■", Target.Offset(, 8).Value)
            見積前提条件 = Replace(見積前提条件, "■■", Target.Offset(, 8).Value)
        End If
        If Target.Offset(, 9).Value = "" Then
            見積もり件名 = Replace(見積もり件名, "▲▲", 変換対象(i, 4))
            見積前提条件 = Replace(見積前提条件, "▲▲", 変換対象(i, 4))
        Else
            見積もり件名 = Replace(見積もり件名, "▲▲", Target.Offset(, 9).Value)
            見積前提条件 = Replace(見積前提条件, "▲▲", Target.Offset(, 9).Value)
        End If
        If Target.Offset(, 10).Value = "" Then
            見積もり件名 = Replace(見積もり件名, "◆◆", 変換対象(i, 5))
            見積前提条件 = Replace(見積前提条件, "◆◆", 変換対象(i, 5))
        Else
            見積もり件名 = Replace(見積もり件名, "◆◆", Target.Offset(, 10).Value)
            見積前提条件 = Replace(見積前提条件, "◆◆", Target.Offset(, 10).Value)
        End If
            
        Set vBook = Workbooks.Add
        ThisWorkbook.Worksheets(WSNAME_FORMAT).Cells.Copy _
            vBook.Worksheets(1).Range("A1")
        With vBook.Worksheets(1)
            EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            For j = 5 To .UsedRange.Rows.Count
                If Len(.Cells(j, 1).Value) = 0 Then
                    EndRow = j - 1
                    Exit For
                End If
            Next
            .Rows(EndRow + 1 & ":" & .Rows.Count).Delete
        End With
        For j = 5 To EndRow
            If vBook.Worksheets(1).Cells(j, 26).Value = 1 Then
                HasKengyoHo = True
                Pub建業法 = "○"
                Exit For
            End If
        Next
        For j = 5 To EndRow
            If vBook.Worksheets(1).Cells(j, 27).Value = 1 Then
                HasShitaukeHo = True
                Pub下請法 = "○"
                Exit For
            End If
        Next
        If HasKengyoHo = True Then
            '            MsgBox "「" & 見積もり件名 & "」の作業開始日を入力してください", vbInformation
            vMsg = "建業法対象データがあります" _
                & vbCrLf & "「" & 見積もり件名 & "」の作業開始日を入力してください"
            MessageBox 0, vMsg, "確認", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            C01_UF01.Show
            開始日 = Pub日付
            '        開始日 = InputBox("作業開始日")
            '            MsgBox "「" & 見積もり件名 & "」の作業終了日を入力してください", vbInformation
            vMsg = "「" & 見積もり件名 & "」の作業終了日を入力してください"
            MessageBox 0, vMsg, "確認", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            C01_UF01.Show
            終了日 = Pub日付
        End If
        With vBook.Worksheets(1)
            '            EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("A3").Value = 1
            .Range("B3").Value = 2
            If Target.Offset(0, 3).Value Like "*回線合算*" Then
                .Range("C3").Value = 4
            Else
                .Range("C3").Value = 2
            End If
            .Range("E3").Value = Pub社員番号 ' Func社員情報取得
            .Range("F3").NumberFormat = "YYYY/MM/DD"
            .Range("F3").Value = Format(DateAdd("M", 見積シート.Range("Y3").Value, Date), "YYYY/MM/DD")
            .Range("G3").Value = 見積シート.Range("Y3").Value
            .Range("H3").Value = 見積シート.Range("AF3").Value
            .Range("I3").Value = 見積シート.Range("AK3").Value
            .Range("J3").Value = 見積シート.Range("AP3").Value
            .Range("K3").Value = 見積もり件名
            .Range("L3").Value = vbNullString   '見積前提条件
            .Range("M3").NumberFormat = "YYYY/MM/DD"
            .Range("M3").Value = Format(開始日, "YYYY/MM/DD")
            .Range("N3").NumberFormat = "YYYY/MM/DD"
            .Range("N3").Value = Format(終了日, "YYYY/MM/DD")
            .Range("O3").Value = ビジネスIT
        End With

        'CSV出力
        vBook.SaveAs Filename:=SavePath & "\" & 見積もり件名 & ".csv", _
            FileFormat:=xlCSV
        
        vBook.Close False

        Open SavePath & "\" & 見積もり件名 & ".txt" For Output As #1

        Print #1, 見積前提条件
 
        Close #1
    Next
    
    Dim 対象ファイル数 As Long
    対象ファイル数 = GetFileCount(SavePath, "csv")
    
    For i = 1 To 対象ファイル数
        If IS_TEST = False Then
            制御文解析 WSNAME_CSVUP
            対象ファイル移動 Pub対象ファイルパス
            
        End If
        実行結果記録
        '        見積番号ファイル保存 SavePath & "\old", Pub件名, Pub見積番号
    Next
    
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Worksheets(WSNAME_LOG)
    
    Dim MsgRange1 As Range
    Dim MsgRange2 As Range
    
    Set MsgRange1 = wsResult.Columns("K").Find("建業法がある時の文言")
    Set MsgRange2 = wsResult.Columns("L").Find("下請法がある時の文言")
    
    If PubHas建業法 Then
        If MsgRange1 Is Nothing Then
        Else
            MessageBox 0, MsgRange1.Offset(1).Value _
                , "確認", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
        End If
    End If
    If PubHas下請法 Then
        If MsgRange2 Is Nothing Then
        Else
            
            MessageBox 0, MsgRange2.Offset(1).Value _
                , "確認", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
        End If
    End If
    ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value = vbNullString
    MsgBox "処理が終了しました", vbInformation
    
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
End Sub

'必須データの入力チェック
Private Function 必須データチェック(ByVal Target As Range _
    , ByVal パターン As Variant) As Boolean
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    Dim temp As String
    If Target.Offset(, 4) = "" Then
        temp = temp & "「件名」"
    End If
    If sh.Range("Y3").Value = "" Then
        temp = temp & "「見積時納期」"
    End If
    If sh.Range("AF3").Value = "" Then
        temp = temp & "「見積納期単位」"
    End If
    If sh.Range("AU3").Value = "" Then
        temp = temp & "「ビジネスIT」"
    End If
    Dim vData As Variant
    vData = モジュール内容取得("営業主担当者コード", "変数名_相対", "A")
    If Trim(vData) = "" Then
        temp = temp & "「営業主担当者コード」"
    End If
    vData = モジュール内容取得("営業主担当者コード", "変数名_相対", "A")

    If Len(temp) > 0 Then
        MsgBox "必須項目の" & temp & "が未入力です" _
            & "確認してください", vbExclamation
        必須データチェック = False
    Else
        必須データチェック = True
    End If
End Function

'Logの記録
Private Sub 実行結果記録()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_LOG)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 5).End(xlUp).Offset(1).Row
        .Cells(LastRow, 5).Value = Format(Now, "YYYY/MM/DD HH:NN")
        .Cells(LastRow, 6).Value = Pub件名
        .Cells(LastRow, 7).Value = Pub見積番号
        .Cells(LastRow, 8).Value = Pub建業法
        .Cells(LastRow, 9).Value = Pub下請法
        If Pub建業法 = "○" Then
            PubHas建業法 = True
        End If
        If Pub下請法 = "○" Then
            PubHas下請法 = True
        End If
    End With
End Sub

'モジュール内容の転記
Private Sub モジュール転記(ByVal パターン As String)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE)
    Dim LastRow As Long
    
    With sh
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 15 To LastRow
        If sh.Cells(i, 3).Value <> "" Then
        モジュール内容入力 sh.Cells(i, 1).Value, "変数名_相対", パターン
        End If
    Next
End Sub

'品目情報の転記
Private Sub 品目情報転記(ByVal 品名 As String _
    , ByVal 対象行 As Long _
    , ByVal 回線合算フラグ As Boolean)
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
    For i = 1 To sh.UsedRange.Rows.Count
        DoEvents
        If sh.Cells(i, 1).Value = 品名 Then
            品目管理表情報入力 "品名", 対象行, i
            品目管理表情報入力 "サービス名", 対象行, i
            品目管理表情報入力 "固定資産（対象・非対象）", 対象行, i
            品目管理表情報入力 "自動契約更新（対象・非対象）", 対象行, i
            品目管理表情報入力 "施策コード", 対象行, i
            品目管理表情報入力 "販売単価", 対象行, i
            品目管理表情報入力 "見積単価_1年目", 対象行, i
            品目管理表情報入力 "見積単価_2年目", 対象行, i
            品目管理表情報入力 "見積単価_3年目", 対象行, i
            品目管理表情報入力 "見積単価_4年目以降", 対象行, i
            品目管理表情報入力 "提供料金区分", 対象行, i
            品目管理表情報入力 "調達方法", 対象行, i
            品目管理表情報入力 "調達単価", 対象行, i
            品目管理表情報入力 "調達単価_1年目", 対象行, i
            品目管理表情報入力 "調達単価_2年目", 対象行, i
            品目管理表情報入力 "調達単価_3年目", 対象行, i
            品目管理表情報入力 "調達単価_4年目以降", 対象行, i
            品目管理表情報入力 "調達料金区分", 対象行, i
            品目管理表情報入力 "社内調達区分", 対象行, i
            品目管理表情報入力 "希望業者", 対象行, i
            品目管理表情報入力 "下請法", 対象行, i
            品目管理表情報入力 "建業法", 対象行, i
            品目管理表情報入力 "TSC保守有", 対象行, i
            品目管理表情報入力 "備考", 対象行, i
            If 回線合算フラグ Then
            品目管理表情報入力 "料金内訳（回線合算のみ使用）", 対象行, i
            Else
            品目管理表情報入力 "料金内訳（回線合算のみ使用）", 対象行, i, True
            End If
            Exit Sub
        End If
    Next
End Sub

'品目管理表情報入力
Private Sub 品目管理表情報入力(ByVal 対象 As String _
    , ByVal 対象行 As Long, ByVal 品目行 As Long _
    , Optional ByVal flg As Boolean = False)
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_HINMOKU)
    
    For i = 4 To sh.UsedRange.Rows.Count
        DoEvents
        If sh.Cells(i, 1).Value = 対象 Then
            
            If flg Then
                ThisWorkbook.Worksheets(sh.Cells(i, 6).Value).Cells(対象行, sh.Cells(i, 8).Value).Value _
                    = ""
            Else
                With ThisWorkbook.Worksheets(sh.Cells(i, 6).Value).Cells(対象行, sh.Cells(i, 8).Value)
                    .NumberFormat = "@"
                    .Value = CStr(ThisWorkbook.Worksheets(WSNAME_HINMOKU).Cells(品目行, sh.Cells(i, 5).Value).Value)
                End With
            End If
            Exit Sub
        End If
    Next
End Sub

Private Sub モジュール内容入力Test()
    モジュール内容入力 "品名2", "変数名_相対", "A"
End Sub
'モジュールの内容を入力
Private Sub モジュール内容入力(ByVal 対象 As String _
    , ByVal 種類 As String, Optional パターン As String)
    Dim 対象範囲 As Range
    Set 対象範囲 = 対象範囲取得(種類)
    Dim vData As Variant
    vData = 対象範囲.Value
    
    Dim 対象行 As Long
    Dim i As Long
    For i = LBound(vData) To UBound(vData)
        DoEvents
        If vData(i, 1) = 対象 Then
            対象行 = i
            Exit For
        End If
    Next
    If 対象行 = 0 Then Exit Sub
    
    If InStr(種類, "相対") > 0 Then
        Dim 基準セル As Range
        Set 基準セル = パターン検索(パターン)
        With ThisWorkbook.Worksheets(vData(対象行, 7)).Cells(vData(対象行, 8), vData(対象行, 9))
            .NumberFormat = "@"
            .Value = CStr(基準セル.Cells(vData(対象行, 5), vData(対象行, 6)).Value)
        End With
    Else
        With ThisWorkbook.Worksheets(vData(対象行, 5)).Cells(vData(対象行, 6), vData(対象行, 7))
            .NumberFormat = "@"
            .Value = CStr(.Worksheets(vData(対象行, 3)).Range(vData(対象行, 4)).Value)
        End With
    End If
End Sub

Private Sub モジュール内容取得Test()
    Debug.Print モジュール内容取得("営業主担当者コード", "変数名_相対", "A")
End Sub
Private Function モジュール内容取得(ByVal 対象 As String _
    , ByVal 種類 As String, Optional パターン As String) As Variant
    Dim 対象範囲 As Range
    Set 対象範囲 = 対象範囲取得(種類)
    Dim vData As Variant
    vData = 対象範囲.Value
    
    Dim 対象行 As Long
    Dim i As Long
    For i = LBound(vData) To UBound(vData)
        DoEvents
        If vData(i, 1) = 対象 Then
            対象行 = i
            Exit For
        End If
    Next
    If 対象行 = 0 Then Exit Function
    
    If InStr(種類, "相対") > 0 Then
        Dim 基準セル As Range
        Set 基準セル = パターン検索(パターン)
        With ThisWorkbook
            モジュール内容取得 = 基準セル.Cells(vData(対象行, 5), vData(対象行, 6)).Value
        End With
    Else
        With ThisWorkbook
            モジュール内容取得 = .Worksheets(vData(対象行, 3)).Range(vData(対象行, 4)).Value
        End With
    End If
End Function

'「変数一覧_モジュール」シートの対象範囲の取得
Private Function 対象範囲取得(ByVal 種類 As String) As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE)
    
    Dim i As Long
    For i = 1 To sh.UsedRange.Rows.Count
        If sh.Cells(i, 1).Value = 種類 Then
            Set 対象範囲取得 = sh.Cells(i, 1).CurrentRegion
            Exit Function
        End If
    Next
    
End Function

Private Sub パターン検索Test()
    Debug.Print パターン検索("A").Address
End Sub
'「品目記入型（モジュール）」シートからパターンを検索
Private Function パターン検索(ByVal パターン As String) As Range
    Dim WS検索対象 As Worksheet
    Set WS検索対象 = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    
    Dim 対象列 As Long
    対象列 = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE).Range("D4").Value

    Dim 最終行 As Long
    With WS検索対象
        最終行 = .Cells(.Rows.Count, 対象列).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 1 To 最終行
        If WS検索対象.Cells(i, 対象列).Value = パターン Then
            Set パターン検索 = WS検索対象.Cells(i, 対象列)
            Exit Function
        End If
    Next
End Function

'「建業法」をチェックする
Public Sub SetKengyoHo(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 26).Value = 1 Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "chk_kengyouhou_" & num Then
                        sh.Cells(j, 8).Value = "○"
                    End If
                Next
            End If
        Next
    End With
End Sub

Private Sub SetFlgTest()
    SetKengyoHo ActiveSheet
    SetShanaiChotatsu ActiveSheet
    SetChotatsukubun ActiveSheet
    SetNextPage
End Sub
'「社内調達」チェック
Public Sub SetShanaiChotatsu(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
'        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 28).Value <> "" Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "sel_chotatsu_hoho_" & num Then
                        sh.Cells(j, 15).Value = .Cells(i, 28).Value
                        sh.Cells(j, 8).Value = "○"
                    End If
                Next
            End If
        Next
    End With
End Sub
'「調達区分」の入力
Public Sub SetChotatsukubun(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
'        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 28).Value <> "" Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "sel_kmo_shanai_choutatsu_" & num Then
                        sh.Cells(j, 15).Value = .Cells(i, 28).Value
                        sh.Cells(j, 8).Value = "○"
                    End If
                Next
            End If
        Next
    End With
End Sub
'次ページへ遷移
Public Sub SetNextPage()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
        .Range(.Cells(20, 7), .Cells(LastRow, 7)).Value = vbNullString
    End With
    
    Dim HasData As Boolean
    Dim TargetRow As Long
    Dim i As Long
    
    '最終データを取得
    For i = LastRow To 20 Step -1
        If sh.Cells(i, 8).Value = "○" Then
            TargetRow = i
            Exit For
        End If
    Next
    
    '最終データより前のページ遷移ボタンはクリック対象
    For i = 20 To TargetRow
        If sh.Cells(i, 14).Value = "btn_next" Then
            sh.Cells(i, 7).Value = "○"
        End If
    Next
    
End Sub
