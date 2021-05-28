VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AlertBox 
   Caption         =   "進捗確認"
   ClientHeight    =   1980
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   4305
   OleObjectBlob   =   "AlertBox.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "AlertBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub cancelBtn_Click()
    cancelFlg = True
End Sub

'@1
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'バツ閉じ対策

    If Not CloseMode = vbFormCode Then
        Cancel = True
    End If
End Sub

'ユーザーフォームを閉じたあと、シートの編集ができなくなるバグ（2013固有？）を回避する
'https://support.microsoft.com/ja-jp/help/2851316
'ユーザーフォームが配置、再配置されると検知

Private Sub UserForm_Layout()
    '静的変数で宣言し、2度目以降はイベント検知しても処理は行わないようにする
    Static fSetModal As Boolean
    If fSetModal = False Then
        fSetModal = True
        'フォームを非表示に
        Me.Hide
        'フォームをモーダルで表示
        Me.Show vbModeless
        
        'フォームを描画させる余裕を作る
        DoEvents

    End If
End Sub


