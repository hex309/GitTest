VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StartTime As Double
 
Property Let ProcessTime(Value As Double)
    StartTime = Value
End Property
 
Property Get ProcessTime() As Double
    ProcessTime = Timer - StartTime
End Property

