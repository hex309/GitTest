Attribute VB_Name = "共通_エラー"
Option Explicit

'-------------------------
'エラー処理
'-------------------------
Public Sub IE不完全操作エラー(ByVal errWin As String)
    
    Dim メッセージ As String
    
    メッセージ = "操作番号: " & Pub操作番号 & vbLf & vbLf & "ErrTitleWin: " & errWin

    MessageBox 0, メッセージ, "オートパイロット停止", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION

    End

End Sub
