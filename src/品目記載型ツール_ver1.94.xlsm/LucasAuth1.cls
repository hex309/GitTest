VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LucasAuth1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim tmpID As String
Dim tmpPW As String

Public Property Let LucasID(ByVal InputID As String)

    tmpID = InputID

End Property

Public Property Let LucasPW(ByVal InputPW As String)

    tmpPW = InputPW

End Property

Public Property Get LucasID() As String

    LucasID = tmpID

End Property

Public Property Get LucasPW() As String

    LucasPW = tmpPW

End Property

