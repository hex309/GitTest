Attribute VB_Name = "M01_HinmokuKinyuTool"
Option Explicit

'「品目管理表」の取得
Public Sub 品目管理表取得()
    MsgBox "品目管理表のデータを取得します", vbInformation
    
    Application.ScreenUpdating = False
'    On Error GoTo ErrHdl
    Dim 品目記入型シート  As String
    Dim 品目管理表セル As String
    Dim vPath As String
    品目記入型シート = 設定取得("▼品目記入型", "品目記入型（モジュール）")
    品目管理表セル = 設定取得("▼品目記入型", "品目管理表ディレクトリパス")
    vPath = ThisWorkbook.Worksheets(品目記入型シート).Range(品目管理表セル).Value
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(vPath) Then
    
    Else
        MsgBox "「品目管理表」ファイルが存在しません" _
            & vbCrLf & "ファイルのパスを確認してください", vbExclamation
        Exit Sub
    End If
    
    Dim WB品目管理表 As Workbook
    Dim WS品目管理表 As Worksheet
    Dim WSコード表 As Worksheet
    
    Set WB品目管理表 = Workbooks.Open(Filename:=vPath, UpdateLinks:=0, ReadOnly:=True)
    Set WS品目管理表 = WB品目管理表.Worksheets("Sheet1")
    Set WSコード表 = WB品目管理表.Worksheets(WSNAME_CODE)
    
    Dim 最終行 As Long
    Dim 最終列 As Long
    Dim 貼付先行 As Long
    
    With WS品目管理表
        最終行 = .Cells(.Rows.Count, 2).End(xlUp).Row
        最終列 = .Cells(5, .Columns.Count).End(xlToLeft).Column
    End With
    Dim i As Long
    Dim WS品目管理シート As Worksheet
    Dim WSコードリスト As Worksheet
    
    Set WS品目管理シート = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
    Set WSコードリスト = ThisWorkbook.Worksheets(WSNAME_CODE)
    
    WS品目管理シート.Range("A1").CurrentRegion.Offset(1).Clear
    WSコードリスト.Cells.Clear
    WSコード表.Cells.Copy Destination:=WSコードリスト.Range("A1")
    
    貼付先行 = 2
    For i = 4 To 最終行
        If WS品目管理表.Cells(i, 2).Value = "品名" Then
            If WS品目管理表.Cells(i + 1, 2).Value <> "" Then
                With WS品目管理表
                    .Range(.Cells(i + 1, 2), .Cells(i + 2, 最終列)).Copy _
                        Destination:=WS品目管理シート.Cells(貼付先行, 1)
                End With
                貼付先行 = 貼付先行 + 1
            End If
        End If
    Next
    
    For i = 2 To 貼付先行 - 1
        With WS品目管理シート
            .Cells(i, 2).NumberFormat = "@"
            .Cells(i, 2).Value = CStr(コード変換(.Cells(1, 2).Value, .Cells(i, 2).Value))
            .Cells(i, 3).NumberFormat = "@"
            .Cells(i, 3).Value = コード変換(.Cells(1, 3).Value, .Cells(i, 3).Value)
            .Cells(i, 4).NumberFormat = "@"
            .Cells(i, 4).Value = コード変換(.Cells(1, 4).Value, .Cells(i, 4).Value)
            .Cells(i, 11).NumberFormat = "@"
            .Cells(i, 11).Value = コード変換(.Cells(1, 11).Value, .Cells(i, 11).Value)
            .Cells(i, 12).NumberFormat = "@"
            .Cells(i, 12).Value = コード変換(.Cells(1, 12).Value, .Cells(i, 12).Value)
            .Cells(i, 13).NumberFormat = "@"
            .Cells(i, 13).Value = コード変換(.Cells(1, 13).Value, .Cells(i, 13).Value)
'            .Cells(i, 17).Value = コード変換(.Cells(1, 17).Value, .Cells(i, 17).Value)
'            .Cells(i, 18).Value = コード変換(.Cells(1, 18).Value, .Cells(i, 18).Value)
'            .Cells(i, 19).Value = コード変換(.Cells(1, 19).Value, .Cells(i, 19).Value)
            .Cells(i, 20).NumberFormat = "@"
            .Cells(i, 20).Value = コード変換(.Cells(1, 20).Value, .Cells(i, 20).Value)
            .Cells(i, 21).NumberFormat = "@"
            .Cells(i, 21).Value = コード変換(.Cells(1, 21).Value, .Cells(i, 21).Value)
            .Cells(i, 22).NumberFormat = "@"
            .Cells(i, 22).Value = コード変換(.Cells(1, 22).Value, .Cells(i, 22).Value)
            .Cells(i, 23).NumberFormat = "@"
            .Cells(i, 23).Value = コード変換(.Cells(1, 23).Value, .Cells(i, 23).Value)
            .Cells(i, 25).NumberFormat = "@"
            .Cells(i, 25).Value = コード変換(.Cells(1, 25).Value, .Cells(i, 25).Value)
            
        End With
    Next
    入力規則設定
    入力規則削除
    
    MsgBox "処理が終了しました", vbInformation
ExitHdl:
    On Error Resume Next
    WB品目管理表.Close False
    On Error GoTo 0
    
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
    Resume ExitHdl
End Sub

'入力規則の設定
Private Sub 入力規則設定()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    
    Dim 最終行 As Long
    With ThisWorkbook.Worksheets(WSNAME_HINMOKU)
        最終行 = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long, j As Long
    For i = 1 To sh.UsedRange.Rows.Count
        For j = 1 To sh.UsedRange.Columns.Count
            If sh.Cells(i, j).Value = "品名" Then
            With sh.Cells(i + 1, j).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop _
                    , Operator:=xlEqual, Formula1:="=品目管理表!$A$2:$A$" & 最終行
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = False
            End With
            End If
        Next
    Next
End Sub

'入力規則の削除
Private Sub 入力規則削除()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
        
    Dim 最終列 As Long
    Dim 最終行 As Long
    With sh
        最終列 = .Cells(1, .Columns.Count).End(xlToLeft).Column
        最終行 = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    With sh.Range(sh.Cells(2, 1), sh.Cells(最終行, 最終列)).Validation
        .Delete
    End With
End Sub

Private Sub コード変換Test()
    Debug.Print コード変換("調達料金区分", "DC年契工事")
End Sub
'文言のコード変換（「コード一覧」シート参照）
Private Function コード変換(ByVal 対象 As String _
    , ByVal 値 As Variant) As Variant
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_CODE)
    
    Dim 対象セル As Range
    Set 対象セル = sh.Rows(1).Find(対象)
    If 対象セル Is Nothing Then
        コード変換 = False
        Exit Function
    End If
    Dim 対象範囲 As Range
    Set 対象範囲 = 対象セル.CurrentRegion
    Dim i As Long
    For i = 2 To 対象範囲.Rows.Count
        If 対象範囲.Cells(i, 1).Value = 値 Then
            コード変換 = 対象範囲.Cells(i, 2).Value
        End If
    Next
End Function



