Attribute VB_Name = "C00_KyotsuFunction"
Option Explicit
Option Private Module

'必須項目確認
Public Function Fnc必須項目確認(ByVal sh As Worksheet) As Boolean
    Dim temp As String
    If sh.Range("D3").Value = "" Then
        temp = temp & sh.Range("D2").Value & ";"
    End If

    If sh.Range("G3").Value = "" Then
        temp = temp & sh.Range("G2").Value & ";"
    End If
    If sh.Range("H3").Value = "" Then
        temp = temp & sh.Range("H2").Value & ";"
    End If
    If sh.Range("K3").Value = "" Then
        temp = temp & sh.Range("K2").Value & ";"
    End If
    If sh.Range("O3").Value = "" Then
        temp = temp & sh.Range("O2").Value & ";"
    End If
    
    If Len(temp) > 0 Then
        temp = Left(temp, Len(temp) - 1)
        MsgBox temp & "は必須項目です。確認してください", vbExclamation
        Fnc必須項目確認 = False
    Else
        Fnc必須項目確認 = True
    End If
    Exit Function
End Function

Private Sub 変数取得Test()
    Debug.Print 変数取得("ブラウザ管理用", "A17", 3, "メインメニュー")
End Sub
'変数取得用
Public Function 変数取得(ByVal 対象シート As String _
    , ByVal 対象テーブル As String _
    , ByVal 対象カラム As Long _
    , ByVal キーワード As String) As Variant
    
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(対象シート)
    Dim Target As Range
    Set Target = sh.Range(対象テーブル).CurrentRegion
    
    For i = 1 To Target.Rows.Count
        If Target.Cells(i, 2).Value = キーワード Then
            変数取得 = Target.Cells(i, 対象カラム).Value
            Exit Function
        End If
    Next
End Function

Private Sub 記号置換テスト()
    MsgBox 記号置換("★★", "【★★】●●拠点　ACL追加対応", "新横浜")
End Sub
'記号を置換する
Public Function 記号置換(ByVal 記号 As String _
    , ByVal 対象文字列 As String _
    , ByVal 置換文字列 As String) As String
    
    Dim temp As String
    temp = Replace(対象文字列, 記号, 置換文字列)
    記号置換 = temp
End Function

'対象のブックを取得
Public Function GetTargetBook(ByVal vPath As String) As Workbook
    Dim vBook As Workbook
    Dim temp As Workbook
    
    On Error GoTo ErrHdl
    For Each temp In Workbooks
        If temp.FullName = vPath Then
            Set GetTargetBook = temp
            Exit Function
        End If
    Next
    Set GetTargetBook = Workbooks.Open(Filename:=vPath)
ExitHdl:
    Exit Function
ErrHdl:
    Resume ExitHdl
End Function

'UsedRange対策
'UsedRangeは、A1が起点になるとは限らないため
Public Sub CorrectUsedRange(ByVal sh As Worksheet _
    , ByVal IsSetting As Boolean)
    If IsSetting Then
        If Trim(sh.Range("A1").Value) = vbNullString Then
            sh.Range("A1").Value = "Dummy"
        End If
    Else
        If sh.Range("A1").Value = "Dummy" Then
            sh.Range("A1").Value = vbNullString
        End If
    End If
End Sub


'ファイルの存在チェックテスト
Private Sub HasTargetFileTest()
    Dim vPath As String
    Dim vFileName As String
    
    vPath = ThisWorkbook.Path
    vFileName = "Test.csv"
    Debug.Print "TRUE:" & HasTargetFile(vPath, vFileName)
    vPath = vPath & "TEST"
    Debug.Print "FALSE:" & HasTargetFile(vPath, vFileName)
End Sub

'ファイルの存在チェック
Public Function HasTargetFile(ByVal vPath As String _
    , ByVal vFileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim TargetPath As String
    TargetPath = vPath & "\" & vFileName
    If fso.FileExists(TargetPath) Then
        HasTargetFile = True
    Else
        HasTargetFile = False
    End If
End Function

'フォルダの存在チェックテスト
Private Sub HasTargetFolderTest()
    Dim vPath As String
    Dim vFileName As String
    
    vPath = ThisWorkbook.Path
    Debug.Print "TRUE:" & HasTargetFolder(vPath)
    vPath = vPath & "TEST"
    Debug.Print "FALSE:" & HasTargetFolder(vPath)
End Sub

'フォルダの存在チェック
Public Function HasTargetFolder(ByVal vPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(vPath) Then
        HasTargetFolder = True
    Else
        HasTargetFolder = False
    End If
End Function

Private Sub GetFileCountTest()
    Debug.Print GetFileCount(ThisWorkbook.Path, "xlsx")
End Sub
'指定した条件のファイルの数をカウント
Public Function GetFileCount(ByVal vPath As String _
    , ByVal vExt As String) As Long
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim num As Long
    Dim temp As Object
    With fso
        For Each temp In .GetFolder(vPath).Files
            If UCase(.GetExtensionName(temp.Path)) _
                = UCase(vExt) Then
                num = num + 1
            End If
        Next
    End With
    GetFileCount = num
End Function


