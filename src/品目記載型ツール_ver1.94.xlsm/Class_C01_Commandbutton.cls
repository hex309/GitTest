VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_C01_Commandbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents CMDBUT As MSForms.CommandButton
Attribute CMDBUT.VB_VarHelpID = -1

Public Sub SetCtrl(new_ctrl As MSForms.CommandButton)

    Set CMDBUT = new_ctrl

End Sub

Private Sub CMDBUT_Click()
    
    If CMDBUT.Caption = "今月" Then Exit Sub

    Dim YYYY As String
    Dim MM As String
    Dim DD As String
    
    Dim 年月日 As String
    Dim 曜日 As String
    
    YYYY = C01_UF01.ComboBox1.Value
    YYYY = Left(YYYY, 4)
    MM = C01_UF01.ComboBox2.Value
    MM = Left(MM, Len(MM) - 1)
    DD = CMDBUT.Caption
    
    年月日 = YYYY & "/" & MM & "/" & DD   '& "日"
'    曜日 = WeekdayName(Weekday(年月日), True)
    
'    年月日 = 年月日 & "(" & 曜日 & ")"

    Dim RC As Long
'    RC = MsgBox(年月日, vbQuestion + vbYesNo, "日付")
    
'    If RC = vbYes Then
'        M01_UF01.TextBox1.Value = 年月日
        Pub日付 = 年月日
        C01_UF01.Hide
'    End If

End Sub


