Attribute VB_Name = "C01_HensuShutoku"
Option Explicit
Option Private Module

'URLなどの値をワークシートから取得
Private Sub 変数取得テスト()
    Debug.Print 設定取得("▼品目記入型", "SSIS URL")
    Debug.Print 設定取得("▼定常コピー型", "SSIS URL")
End Sub
'URLなどの値をワークシートから取得
Public Function 設定取得(ByVal 対象 As String _
    , ByVal 名称 As String) As Variant
    Dim sh設定 As Worksheet
    Set sh設定 = ThisWorkbook.Worksheets(WSNAME_CONFIG)
    
    Dim 対象列 As Range
    With sh設定
        Set 対象列 = .Rows(1).Find(対象)
    End With
    
    Dim oResult As Range
    With sh設定
        Set oResult = .Columns(対象列.Column).Find(名称)
    End With
    If oResult Is Nothing Then
        設定取得 = vbNullString
    Else
        設定取得 = oResult.Offset(0, 1).Value
    End If
End Function
