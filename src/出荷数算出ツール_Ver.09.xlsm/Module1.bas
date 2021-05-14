Attribute VB_Name = "Module1"
Option Explicit

Public Sub GetCurrentData()
    '�N���̃`�F�b�N
    Dim TargetYear As Variant
    TargetYear = GetTargetYear
    If TargetYear = False Then
        MsgBox "���i�䒠�i�����j�̐擪���N�x��\���Ă��܂���B�������I�����܂�", vbInformation
        Exit Sub
    End If
    
    '�f�[�^�̎擾
    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    Dim TargetPath As String
    Dim DataSheet As Worksheet
    Dim TempSheet As Worksheet
    
    '    On Error GoTo ErrHdl
    Application.ScreenUpdating = False
    
    GetSupplier
    
    Set DataSheet = ThisWorkbook.Worksheets("Data")
    DataSheet.Cells.Clear
    With DataSheet
        .Range("A1").Value = "�i��"
        .Range("B1").Value = "�O���c"
        .Range("C1").Value = "������"
        .Range("J1").Value = "�i��  "
        .Range("K1").Value = "�O���c"
        .Range("L1").Value = "������"
        .Range("N1").Value = "�i�ڔԍ�"
        .Range("O1").Value = "�������א���"
        .Range("Q1").Value = "�����ԍ�"
    End With
    Set TempSheet = ThisWorkbook.Worksheets("Temp")
    TempSheet.Cells.Clear
    
    TargetPath = FncGetFileSetting(uexlProductBook)
    Set TargetBook = Workbooks.Open(Filename:=TargetPath _
        , UpdateLinks:=0, ReadOnly:=True)
    Set TargetSheet = TargetBook.Worksheets("X_���i�䒠")
    With TargetSheet
        If .AutoFilterMode Then
            .Range("A1").CurrentRegion.AutoFilter
        End If
        With .Range("A1").CurrentRegion
            .AutoFilter Field:=2, Criteria1:="���i"
            .AutoFilter Field:=12, Criteria1:=">0"
            .Columns(1).Copy Destination:=TempSheet.Columns(1)
            .Columns(12).Copy Destination:=TempSheet.Columns(2)
            .Columns(6).Copy Destination:=TempSheet.Columns(3) '������
        End With
    End With
    TargetBook.Close False
    '�W�v
    Dim DataArray As Variant
    Dim Target As Range
    
    DataArray = TempSheet.Range("A1").CurrentRegion.Value
    Set Target = DataSheet.Range("A2")
    SetData DataArray, Target

    TargetPath = FncGetFileSetting(uexlPreProductBook)
    Set TargetBook = Workbooks.Open(Filename:=TargetPath _
        , UpdateLinks:=0, ReadOnly:=True)
    Set TargetSheet = TargetBook.Worksheets("X_���i�䒠")
    With TargetSheet
        If .AutoFilterMode Then
            .Range("A1").CurrentRegion.AutoFilter
        End If
        With .Range("A1").CurrentRegion
            .AutoFilter Field:=2, Criteria1:="���i"
            .AutoFilter Field:=12, Criteria1:=">0"
            .Columns(1).Copy Destination:=TempSheet.Columns(1)
            .Columns(12).Copy Destination:=TempSheet.Columns(2)
            .Columns(6).Copy Destination:=TempSheet.Columns(3) '������
        End With
    End With
    TargetBook.Close False

    DataArray = TempSheet.Range("A1").CurrentRegion.Value
    Set Target = DataSheet.Range("J2")
    SetData DataArray, Target
    
    TargetPath = FncGetFileSetting(uexlOrderArrival)
    Set TargetBook = Workbooks.Open(Filename:=TargetPath _
        , UpdateLinks:=0, ReadOnly:=True)
    Set TargetSheet = TargetBook.Worksheets(1)
    With TargetSheet.Range("A1").CurrentRegion
        .Columns(4).Copy Destination:=TempSheet.Columns(1)
        .Columns(5).Copy Destination:=TempSheet.Columns(2)
        .Columns(3).Copy Destination:=TempSheet.Columns(3)
    End With
    TargetBook.Close False

    DataArray = TempSheet.Range("A1").CurrentRegion.Value
    Set Target = DataSheet.Range("N2")
    SetData DataArray, Target
    
    '�v�Z��
    With DataSheet.Range("A1").CurrentRegion
        .Offset(1, .Columns.Count).Resize(.Rows.Count - 1, 1).Formula _
            = "=VLOOKUP(RC[-3],C[6]:C[7],2,FALSE)"
        .Offset(1, .Columns.Count + 1).Resize(.Rows.Count - 1, 1).Formula _
            = "=VLOOKUP(RC[-4],C[9]:C[11],2,FALSE)"
        .Offset(1, .Columns.Count + 2).Resize(.Rows.Count - 1, 1).Formula _
            = "=VLOOKUP(RC[-3],�d����!C[-5]:C[-4],2,FALSE)"
        '        .Offset(1, .Columns.Count + 3).Resize(.Rows.Count - 1, 1).Formula _
                 = "=VLOOKUP(RC[-3],�d����!C[-5]:C[-4],2,FALSE)"
    End With
    With DataSheet.Range("A1").CurrentRegion
        .Value = .Value
        .Replace "#N/A", ""
        .Offset(1, .Columns.Count).Resize(.Rows.Count - 1, 1).Formula _
            = "=RC[-5]+RC[-2]-RC[-3]"
    End With
    
    '�f�[�^�V�[�g�̊����f�[�^���`�F�b�N
    If HasData(TargetYear) Then
        '�f�[�^������ꍇ�͍폜
        DeleteData TargetYear
    End If
    
    '���N�x�V�[�g�̃f�[�^�`�F�b�N
    '    If AddAnnualData(TargetYear) Then
    '        '�����̏ꍇ�͏㏑������̂Ńf�[�^�N���A
    '        ClearData TargetYear
    '    End If
    
    Dim TotalData As Worksheet
    Set TotalData = ThisWorkbook.Worksheets("�f�[�^")
    Dim TargetRow As Long
    With ThisWorkbook.Worksheets("�f�[�^")
        TargetRow = .Cells(.Rows.Count, 1).End(xlUp).Offset(1).Row
    End With
    
    Dim LastRow As Long
    Dim vData As Variant
    
    With DataSheet.Range("A1").CurrentRegion
        .Offset(1).Resize(.Rows.Count - 1, 1).Copy TotalData.Cells(TargetRow, 1)
        .Offset(1, 5).Resize(.Rows.Count - 1, 1).Copy TotalData.Cells(TargetRow, 2)
        vData = .Offset(1, 6).Resize(.Rows.Count - 1, 1).Value
        LastRow = .Rows.Count - 1
    End With
    With TotalData
        .Range(.Cells(TargetRow, 3), .Cells(TargetRow + LastRow - 1, 3)).Value _
            = vData
        .Range(.Cells(TargetRow, 4), .Cells(TargetRow + LastRow - 1, 4)).Value _
            = TargetYear
        .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    End With
    
    '�s�{�b�g�e�[�u���̍X�V
    UpdatePivot TotalData
                                    
    '�]�L
    CopyPivot
    
    '�N�ō쐬
    If AddAnnualData(TargetYear) = False Then
        Exit Sub
    End If
    
    '�ۑ�
    SaveData
    
ExitHdl:
    MsgBox "�������I�����܂���", vbInformation
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
    Resume ExitHdl
End Sub

'���f�[�^���u�i�ԁv���ƂɏW�v
Private Sub SetData(ByVal DataArray As Variant, ByVal Target As Range)
    Dim oProduct As ProductItems
    Set oProduct = New ProductItems
    Dim i As Long
    For i = LBound(DataArray) + 1 To UBound(DataArray)
        oProduct.Add DataArray(i, 1), DataArray(i, 2), DataArray(i, 3)
    Next
    Dim temp As Variant
    temp = oProduct.GetAllData
    Target.Resize(UBound(temp), UBound(temp, 2)).Value = temp
End Sub

'�Ώۃt�@�C���̐擪4�������`�F�b�N
'�Ώ۔N�x���擾
Private Function GetTargetYear() As Variant
    Dim TargetPath As String
    Dim TargetFileName As String
    Dim tempYear As Long
    TargetPath = FncGetFileSetting(uexlPreProductBook)
    
    TargetFileName = GetFileName(TargetPath)
    tempYear = Left(TargetFileName, 4)
    
    If IsNumeric(tempYear) Then
        GetTargetYear = tempYear
    Else
        GetTargetYear = False
    End If
    
End Function

Private Sub GetFileNameTest()
    Debug.Print GetFileName(ThisWorkbook.FullName)
End Sub

Private Function GetFileName(ByVal FullName As String) As Variant
    Dim pos As Long
    
    pos = InStrRev(FullName, "\")
    If pos = 0 Then
        GetFileName = False
    Else
        GetFileName = Mid(FullName, pos + 1)
    End If
End Function

Private Function HasData(ByVal TargetYear As Long) As Boolean
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�f�[�^")
    Dim i As Long
    
    With sh
        For i = 1 To sh.Range("A1").CurrentRegion.Rows.Count
            If .Cells(i, 4).Value = TargetYear Then
                HasData = True
                Exit Function
            End If
        Next
    End With
    HasData = False
End Function

Private Sub DeleteDataTest()
    DeleteData 2014
End Sub

Private Sub DeleteData(ByVal TargetYear As Long)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�f�[�^")
    If sh.AutoFilterMode Then
        sh.Range("A1").CurrentRegion.AutoFilter
    End If
    
    sh.Range("A1").CurrentRegion.AutoFilter Field:=4, Criteria1:=TargetYear
    With sh.AutoFilter.Range
        '�f�[�^���Ȃ��ꍇ�͏������s��Ȃ�
        If .SpecialCells(xlCellTypeVisible).Rows.Count > 1 Then
            .Offset(1).Resize(.Rows.Count - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        End If
    End With
    sh.Range("A1").CurrentRegion.AutoFilter
End Sub

Private Sub SortYear()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("�f�[�^")
    sh.Range("A1").CurrentRegion.Sort _
    Key1:=Range("D1"), Order1:=xlAscending, Header:=xlYes
End Sub

Private Sub UpdatePivot(ByVal TotalData As Worksheet)
    Dim vName As String
    Dim vAddress As String
    vName = TotalData.Name
    vAddress = TotalData.Range("A1").CurrentRegion.Address
    ThisWorkbook.Worksheets("Pivot").PivotTables(1).ChangePivotCache _
        ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
        SourceData:=vName & "!" & vAddress)
End Sub

Private Sub CopyPivot()
    With ThisWorkbook.Worksheets("����")
        .Cells.Clear
        ThisWorkbook.Worksheets("Pivot").Range("A3").CurrentRegion.Copy
        .Range("A1").PasteSpecial xlValues
        .Rows(1).Delete
        With .Range("A1").CurrentRegion
            .Borders.LineStyle = xlContinuous
            .Columns.AutoFit
        End With
        With .Range("A1").CurrentRegion.Rows(1)
            .Interior.Color = RGB(217, 217, 217)
            .Font.Bold = True
        End With
        
    End With
End Sub

Private Sub AddAnnualDataTest()
    Debug.Print AddAnnualData(2022)
End Sub

Private Function AddAnnualData(ByVal vYear As Long) As Boolean
    Dim DataSheet As Worksheet
    Dim sh As Worksheet
    
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name Like "*�N��" Then
            Set DataSheet = sh
            Exit For
        End If
    Next
    If DataSheet Is Nothing Then
        MsgBox "�u�����N�Łv�̃��[�N�V�[�g���L��܂���B�������I�����܂�", vbExclamation
        AddAnnualData = False
        Exit Function
    End If
    
    Dim vCol As Long
    Dim i As Long
    
    '    vYear = Format(Date, "YYYY") - 1
    Dim IsNewYear As Boolean
    IsNewYear = True
    With DataSheet.Range("A1").CurrentRegion
        '�㏑�����邩�̃`�F�b�N
        For i = 1 To .Columns.Count
            If .Cells(1, i).Value = vYear Then
                vCol = i
                IsNewYear = False
                Exit For
            End If
        Next
        '�V�N�x�̏ꍇ
        If vCol = 0 Then
            For i = 1 To .Columns.Count
                If .Cells(1, i).Value = vYear - 1 Then
                    vCol = i
                    vCol = vCol + 1
                    DataSheet.Columns(vCol).Insert
                    IsNewYear = True
                    Exit For
                End If
            Next
        End If
        If vCol = 0 Then
            MsgBox "�Ώ۔N�x�̗񂪂���܂���B�m�F���Ă�������", vbExclamation
            AddAnnualData = False
            Exit Function
        End If
    End With
    '�V���ȕi�ڂ̒ǉ�
    If AddNewData = False Then
        Exit Function
    End If
        
    With DataSheet.Range("A1").CurrentRegion
        Dim TargetCol As Long
        Dim vOffsetNum1 As Long
        Dim vOffsetNum2 As Long
        If IsNewYear Then
            .Cells(1, vCol).Value = vYear
            DataSheet.Name = vYear + 1 & "�N��"
        Else
            DataSheet.Range("A1").CurrentRegion.Resize(, 1).Offset(1, vCol - 1).ClearContents
        End If
        
        '�݌ɐ�
        TargetCol = GetTargetCol(DataSheet, vYear)
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=IFERROR(VLOOKUP(RC[" & -TargetCol + 1 _
                & "],Data!C[" & -TargetCol + 1 & "]:C[" & -TargetCol + 7 & "],7,FALSE),"""")"
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value _
                = .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value
        End If

        '���v
        TargetCol = GetTargetCol(DataSheet, "���v")
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=SUM(RC[" & -TargetCol + 3 & "]:RC[-1])"
            '                            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value _
                                         = .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value
        End If
        '4�N����
        TargetCol = GetTargetCol(DataSheet, "�@4�N����")
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=AVERAGE(RC[-2]:RC[-5])"
        End If
            
        '�@/4Q
        TargetCol = GetTargetCol(DataSheet, "�@/4Q")
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=RC[-1]/4"
        End If
        '3�����݌�
        TargetCol = GetTargetCol(DataSheet, "3�����݌�")
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=IFERROR(VLOOKUP(RC[" & -TargetCol + 1 _
                & "],Data!C[" & -TargetCol + 10 & "]:C[" & -TargetCol + 11 & "],2,FALSE),"""")"
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value _
                = .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Value
        End If
        TargetCol = GetTargetCol(DataSheet, "����")
        If TargetCol > 0 Then
            .Range(.Cells(2, TargetCol), .Cells(.Rows.Count, TargetCol)).Formula _
                = "=IF(RC[-2]>RC[-1],""��"",""�~"")"
        End If
        
    End With
    AddAnnualData = True
End Function

Private Sub AddNewDataTest()
    Debug.Print AddNewData
End Sub
Private Function AddNewData() As Boolean
    Dim OriginRange As Range
    Dim OriginData As Variant
    Dim TargetSheet As Worksheet
    Dim TargetRange As Range
    Dim TargetData As Variant
    Set OriginRange = ThisWorkbook.Worksheets("Data").Range("A1").CurrentRegion
    OriginData = OriginRange.Value
    If VarType(SetTargetSheet) = vbBoolean Then
        AddNewData = False
        Exit Function
    Else
        Set TargetSheet = SetTargetSheet
    End If
    Set TargetRange = SetTargetSheet.Range("A1").CurrentRegion
    TargetData = TargetRange.Value
    
    '�V�K�f�[�^��z��ɒǉ�
    Dim NewData() As Variant
    Dim HasData As Boolean
    Dim i As Long, j As Long
    Dim Num As Long
    
    For i = 1 To UBound(OriginData)
        HasData = False
        For j = 1 To UBound(TargetData)
            If OriginData(i, 1) = TargetData(j, 1) Then
                HasData = True
                Exit For
            End If
        Next
        If HasData Then
        
        Else
            Num = Num + 1
            ReDim Preserve NewData(1 To 2, 1 To Num)
            NewData(1, Num) = OriginData(i, 1)
            NewData(2, Num) = OriginData(i, 6)
        End If
    Next
    
    Dim vRow As Long
    If Num > 0 Then
        vRow = TargetSheet.Cells(1, 1).End(xlDown).Offset(1).Row
        With TargetSheet
            .Range(.Cells(vRow, 1), .Cells(vRow + Num - 1, 2)).Value _
                = Application.WorksheetFunction.Transpose(NewData)
        End With
    End If
    AddNewData = True
    
    SetTargetSheet.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
End Function

Private Function GetTargetCol(ByVal sh As Worksheet, ByVal Target As String) As Long
    Dim i As Long
    With sh.Range("A1").CurrentRegion
        For i = 1 To .Columns.Count
            If .Cells(1, i).Value = Target Then
                GetTargetCol = i
                Exit Function
            End If
        Next
    End With
End Function

Private Function SetTargetSheet() As Variant
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name Like "*�N��" Then
            Set SetTargetSheet = sh
            Exit Function
        End If
    Next
    SetTargetSheet = False
End Function

Private Sub SaveData()
    Dim vPath As String
    vPath = ThisWorkbook.Worksheets("���f�[�^�Ǎ�").Range(SAVE_FOLDER).Value
    
    If Len(Dir(vPath, vbDirectory)) = 0 Then
        vPath = ThisWorkbook.Path
    End If
    
    Dim wb As Workbook
    Set wb = Workbooks.Add
    Dim DataArray As Variant
    '4�̃V�[�g��ۑ�����
    Dim FYSheetName As String
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name Like "*�N��" Then
            FYSheetName = sh.Name
            Exit For
        End If
    Next
    DataArray = Array("����", "Pivot", "�f�[�^", FYSheetName)
    ThisWorkbook.Worksheets(DataArray).Copy After:=wb.Worksheets(1)
    Application.DisplayAlerts = False
    wb.Worksheets(1).Delete
        
    wb.SaveAs vPath & "\" & Format(Date, "YYYY") - 1 & "�N�x_�o�׎���.xlsx"
    Application.DisplayAlerts = True
End Sub

Public Sub GetSupplier()
    Dim TargetPath As String
    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
'    TargetPath = FncGetFileSetting(uexlSupplierBook)
    TargetPath = ThisWorkbook.Worksheets("���f�[�^�Ǎ�").Range("D30").Value
    If Len(TargetPath) > 0 Then
        Set TargetBook = Workbooks.Open(Filename:=TargetPath _
            , UpdateLinks:=0, ReadOnly:=True)
        Set TargetSheet = TargetBook.Worksheets(1)
    
        ThisWorkbook.Worksheets("�d����").Cells.Clear
        TargetSheet.Range("A1").CurrentRegion.Copy _
            Destination:=ThisWorkbook.Worksheets("�d����").Range("A1")
        TargetBook.Close False
    End If
End Sub



