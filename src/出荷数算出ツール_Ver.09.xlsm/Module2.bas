Attribute VB_Name = "Module2"
Option Explicit

Private Sub Test()
    Dim DataArray As Variant
    DataArray = ThisWorkbook.Worksheets("Temp").Range("A1").CurrentRegion.Value
    Dim Target As Range
    Set Target = ThisWorkbook.Worksheets("Data").Range("A2")
    
    SetData DataArray, Target
End Sub

Private Sub SetData(ByVal DataArray As Variant, ByVal Target As Range)
    Dim oProduct As ProductItems
    Set oProduct = New ProductItems
    Dim i As Long
    For i = LBound(DataArray) + 1 To UBound(DataArray)
        oProduct.Add DataArray(i, 1), DataArray(i, 2)
    Next
    Dim temp As Variant
    temp = oProduct.GetAllData
    Target.Resize(UBound(temp), UBound(temp, 2)).Value = temp
End Sub
