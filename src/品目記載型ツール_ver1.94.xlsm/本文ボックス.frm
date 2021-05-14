VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 本文ボックス 
   Caption         =   "見積前提条件"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   OleObjectBlob   =   "本文ボックス.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "本文ボックス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Modシート名 As String

Private Sub Userform_initialize()

    Modシート名 = ThisWorkbook.ActiveSheet.Name

    TextBox1.Value = Replace(ThisWorkbook.Sheets(Modシート名).Range(Pub本文アドレス).Value, vbCr, "")
    
End Sub

'-----------------------------------------
' 「閉じる」ボタン
'-----------------------------------------
Private Sub CommandButton1_Click()
        
    ThisWorkbook.Sheets(Modシート名).Range(Pub本文アドレス).Value = TextBox1.Value
    
    Unload Me

End Sub
