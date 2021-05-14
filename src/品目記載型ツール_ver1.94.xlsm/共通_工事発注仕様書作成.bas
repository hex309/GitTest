Attribute VB_Name = "共通_工事発注仕様書作成"
Option Explicit

Dim Modアドオンフルパス As String
Dim Modアドオン実行モジュール As String
Dim Modツールフルパス As String
    
Sub 共通_工事発注仕様書作成モジュール()
    
    Call アドオン初期設定
    Call アドオン実行

End Sub

Sub アドオン初期設定()

    Modアドオンフルパス = ThisWorkbook.Path & "\" & "【add-on】SSIS自動登録.xlam"
    Modアドオン実行モジュール = "メイン_工事発注仕様書作成.メイン_工事発注仕様書作成モジュール"
    
    Modツールフルパス = ThisWorkbook.FullName
    
End Sub

Sub アドオン実行()
    
    Dim アドオンフルパス As String
    Dim アドオン実行モジュール As String
    Dim アドオン途中停止フラグ As Boolean
    Dim ツールフルパス As String
   
    '-----------------------------------------------
    ' アドオンファイルは、同一階層に存在すること。
    '-----------------------------------------------
    アドオンフルパス = Modアドオンフルパス
    アドオン実行モジュール = Modアドオン実行モジュール
    ツールフルパス = Modツールフルパス
    
    '-----------------------------------------------
    ' アドオン存在確認
    '-----------------------------------------------
    If Dir(アドオンフルパス) = "" Then
        MsgBox アドオンフルパス & "が存在しないため中止します", vbExclamation
        End
    End If
    
    '-----------------------------------------------
    ' アドオン実行
    '-----------------------------------------------
    Dim strJoin As String
    strJoin = "'" & アドオンフルパス & "'!" & アドオン実行モジュール
    
    Application.Run strJoin, ツールフルパス, Pub見積シート名, Pub工事発注仕様書, Pub店舗コード, Pub見積登録件名

End Sub


