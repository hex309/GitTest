Attribute VB_Name = "例外_見積出力"
Option Explicit

Sub 例外_見積出力モジュール()

'    Call 見積確認書情報登録アップロード用CSVファイル出力
    Call 例外_見積確認書情報登録_割込制御

End Sub


'Sub 見積確認書情報登録アップロード用CSVファイル出力()
'
'    Dim シート名    As String
'
'    Call 共通_見積登録アップロード用.見積確認書情報登録アップロード用CSVファイル作成(Pub見積シート名)
'
'End Sub

Sub 例外_見積確認書情報登録_割込制御()
   
    ThisWorkbook.Sheets(WSNAME_WARIKOMI).Range("H20:H100").ClearContents
    
    Dim buf As String
    Dim tmp As Variant
    
    Open ThisWorkbook.Path & "\★アップロード用ファイル\見積確認書情報登録(アップロード用).csv" For Input As #1
    
    Dim i As Long: i = 1
    Dim z As Long: z = 0
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")

        If i >= 6 Then
            If tmp(25) = 1 Then ThisWorkbook.Sheets(WSNAME_WARIKOMI).Cells(14 + i, 8).Value = "○"  '下請法対象
            If tmp(26) = 1 Then ThisWorkbook.Sheets(WSNAME_WARIKOMI).Cells(14 + i + 1, 8).Value = "○"  '建業法対象
            i = i + 1
        End If
        If 14 + i = 39 Then i = i + 2
        i = i + 1
    Loop
    
    Close #1

End Sub

