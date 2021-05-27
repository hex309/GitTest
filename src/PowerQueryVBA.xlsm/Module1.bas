Attribute VB_Name = "Module1"
Option Explicit

'https://qiita.com/hazigin/items/e37f2488d5f33b50240e

'------------------------------------------------
'  引  数：arg_sheet_name    対象シート
'  戻り値：生成されたクエリ名
'------------------------------------------------
Function ConvertTableToQuery(arg_sheet_name As String) As String
On Error GoTo HandleErr
    Dim Table_Name As String: Table_Name = arg_sheet_name & "_Table"
    Dim SheetName As String: SheetName = arg_sheet_name

    If IsQuerrExist(Table_Name) = True Then
        MsgBox Table_Name & ":Already Exist"
        ConvertTableToQuery = vbNullString
        Exit Function
    End If

    ThisWorkbook.Sheets(SheetName).ListObjects.Add(xlSrcRange, ThisWorkbook.Sheets(SheetName).Range("$A$1:$" & ConvNumToAlphabet(CountColumn(ThisWorkbook.Sheets(SheetName))) & "$" & CountRow(ThisWorkbook.Sheets(SheetName))), , xlYes).Name = Table_Name
    ThisWorkbook.Queries.Add Name:=Table_Name, Formula:="let Source = Excel.CurrentWorkbook(){[Name=""" & Table_Name & """]}[Content],  MODIED_TYPE = Table.TransformColumnTypes(Source,{" & GetShemaString(ThisWorkbook.Sheets(SheetName)) & "}) , AddIndex = Table.AddIndexColumn(MODIED_TYPE, ""Index_a"", 0, 1, Int64.Type) in AddIndex"

    ConvertTableToQuery = Table_Name
    Exit Function
HandleErr:
    MsgBox "Error ConvertTableToQuery:" & Err.Number & vbCrLf & "Description:" & Err.Description
    ConvertTableToQuery = vbNullString
End Function


'------------------------------------------------
'  引  数：対象シート
'  戻り値：カラム一覧
'------------------------------------------------
Function GetShemaString(SheetName As Worksheet)
    Dim ColumnArr() As String
    Dim i As Long
    Dim rec As String: rec = vbNullString
    Dim t_maxcol As Long
    t_maxcol = CountColumn(SheetName)


    For i = 1 To t_maxcol
        ReDim Preserve ColumnArr(i - 1)
        ColumnArr(i - 1) = "{""" & SheetName.Cells(1, i).Value & """, type text}"
    Next i

    For i = 0 To UBound(ColumnArr)
        rec = rec & ColumnArr(i) & ","
    Next i

    rec = CutRight(rec, 1)

    GetShemaString = rec

End Function

'------------------------------------------------
'  引  数：
'  arg_LeftQuery    結合の基準となるテーブル名
'  argRightQuery    他方のテーブル
'  Key_col    結合条件のカラム
'  MergeQueryName    生成するクエリ名
'  戻り値：True 成功 / False 失敗
'------------------------------------------------
Function DistMergeQuery(arg_LeftQuery As String, argRightQuery As String, Key_col As String, MergeQueryName As String) As Boolean
On Error GoTo HandleErr

    ActiveWorkbook.Queries.Add Name:=MergeQueryName, Formula:="let Source = Table.NestedJoin(" & arg_LeftQuery & ", {""" & Key_col & """}, " & argRightQuery & ", {""" & Key_col & """}, """ & argRightQuery & """, JoinKind.LeftOuter),  Merged_Table = Table.ExpandTableColumn(Source, """ & argRightQuery & """, {" & GetShemaStringForMerge(ThisWorkbook.Sheets(Left(argRightQuery, Len(argRightQuery) - 6))) & "}, {" & GetShemaStringForMerge2(ThisWorkbook.Sheets(Left(argRightQuery, Len(argRightQuery) - 6))) & "}), AddIndex = Table.AddIndexColumn(Merged_Table , """ & MergeQueryName & "_Index" & """, 0, 1, Int64.Type) in AddIndex"
    DistMergeQuery = MergeQueryName

    If DistQueryToSheet(ThisWorkbook.Sheets("WorkQueryDist"), MergeQueryName, "$A$1") = False Then
        DistMergeQuery = False
        Exit Function
    End If
    DistMergeQuery = True
    Exit Function
HandleErr:
    MsgBox "Error DistMergeQuery:" & Err.Number & vbCrLf & "Description:" & Err.Description
    DistMergeQuery = False
End Function


'------------------------------------------------
'  引  数：対象シート
'  戻り値：カラム一覧
'------------------------------------------------
Function GetShemaStringForMerge(SheetName As Worksheet) As String
    Dim ColumnArr() As String
    Dim i As Long
    Dim rec As String: rec = vbNullString
    Dim t_maxcol As Long
    t_maxcol = CountColumn(SheetName)

    For i = 1 To t_maxcol
        ReDim Preserve ColumnArr(i - 1)
        ColumnArr(i - 1) = """" & SheetName.Cells(1, i).Value & """"
    Next i

    For i = 0 To UBound(ColumnArr)
        rec = rec & ColumnArr(i) & ","
    Next i

    rec = CutRight(rec, 1)

    GetShemaStringForMerge = rec

End Function


'------------------------------------------------
'  引  数：対象シート
'  戻り値：カラム一覧
'------------------------------------------------
Function GetShemaStringForMerge2(SheetName As Worksheet) As String
    Dim ColumnArr() As String
    Dim i As Long
    Dim rec As String: rec = vbNullString
    Dim t_maxcol As Long
    t_maxcol = CountColumn(SheetName)

    For i = 1 To t_maxcol
        ReDim Preserve ColumnArr(i - 1)
        ColumnArr(i - 1) = """" & SheetName.Name & "_Table." & SheetName.Cells(1, i).Value & """"
    Next i

    For i = 0 To UBound(ColumnArr)
        rec = rec & ColumnArr(i) & ","
    Next i

    rec = CutRight(rec, 1)

    GetShemaStringForMerge2 = rec

End Function


'------------------------------------------------
'  引  数：対象シート
'  戻り値：True 成功 / False 失敗
'------------------------------------------------
Function DistQueryToSheet(sheet As Worksheet, queryname As String, rangestr As String) As Boolean
On Error GoTo HandleErr
    With sheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryname & ";Extended Properties=""""" _
        , Destination:=sheet.Range(rangestr)).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryname & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = queryname
        .Refresh BackgroundQuery:=False
    End With
    DistQueryToSheet = True
    Exit Function
HandleErr:
    MsgBox "Error DistQueryToSheet:" & Err.Number & vbCrLf & "Description:" & Err.Description
    DistQueryToSheet = False
End Function


'------------------------------------------------
'  引  数：対象シート
'------------------------------------------------
Sub TableUnlist(sheet As Worksheet)
    On Error Resume Next
    Dim ls As ListObject

    For Each ls In sheet.ListObjects
        ls.TableStyle = ""
        ls.Unlist
    Next ls
End Sub


'------------------------------------------------
'  引  数：検査対象クエリ名
'  戻り値：True 存在する / False 存在しない
'------------------------------------------------
Function IsQuerrExist(qname As String) As Boolean
    Dim rec As Boolean: rec = False
    Dim wq As WorkbookQuery
    For Each wq In ThisWorkbook.Queries
          If InStr(1, wq.Name, qname) > 0 Then
            rec = True
          End If
    Next wq
    IsQuerrExist = rec
End Function
