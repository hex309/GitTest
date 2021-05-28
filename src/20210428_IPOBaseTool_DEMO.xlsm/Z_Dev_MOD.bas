Attribute VB_Name = "Z_Dev_MOD"
Option Explicit

Sub Msg_Unvis()
    With Sheets("インポート")
        Application.ScreenUpdating = True
        .Unprotect
        .Shapes("Fil_1").Visible = False
        .Shapes("Gr_1").Visible = False
    End With
    
End Sub
