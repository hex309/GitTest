VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StopButton 
   Caption         =   "オートパイロット実行中"
   ClientHeight    =   1425
   ClientLeft      =   16620
   ClientTop       =   5460
   ClientWidth     =   2625
   OleObjectBlob   =   "StopButton.frx":0000
End
Attribute VB_Name = "StopButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()

    fStop = True

End Sub
