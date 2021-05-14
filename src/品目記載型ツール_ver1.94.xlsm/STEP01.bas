Attribute VB_Name = "STEP01"
Option Explicit

Sub STEP01モジュール()

    Call 事前準備処理

End Sub

Sub 事前準備処理()

    Call 共通_IE制御.ieExistCheck
'    Call 共通_フォルダ作成.各種フォルダ作成
    
    If Pubオートパイロット番号 = "" Then MsgBox "オートパイロット番号がありません。終了します。", vbExclamation: End
    If Pub制御シート名 = "" Then MsgBox "制御シート名がありません。終了します。", vbExclamation: End
    If Pub見積シート名 = "" Then MsgBox "見積シート名がありません。終了します。", vbExclamation: End
    If Pub見積登録件名 = "" Then MsgBox "▼件名がありません。終了します。", vbExclamation: End
    If Pub工事発注仕様書 = "" Then MsgBox "▼工事発注仕様書が選択されていません。終了します。", vbExclamation: End
    
    If Pub工事発注仕様書 = "建業法" Then
        If Pub工期FROM = "" Then MsgBox "建業法対象は、▼工期FROMが必要です。終了します。", vbExclamation: End
        If Pub工期TO = "" Then MsgBox "建業法対象は、▼工期TOが必要です。終了します。", vbExclamation: End
        If Pub主任者コード = "" Then MsgBox "建業法対象は、▼主任者コードが必要です。終了します。", vbExclamation: End
    End If
    
    If Pub工事発注仕様書 <> "なし" Then
        If Pub店舗コード = "" Then
            MsgBox "▼店舗コードがありません。発注仕様書作成の際に必要です。終了します。", vbExclamation: End
        End If
        Call 共通_工事発注仕様書作成.共通_工事発注仕様書作成モジュール
    End If
    
End Sub


