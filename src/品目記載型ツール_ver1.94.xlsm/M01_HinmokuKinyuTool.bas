Attribute VB_Name = "M01_HinmokuKinyuTool"
Option Explicit

'�u�i�ڊǗ��\�v�̎擾
Public Sub �i�ڊǗ��\�擾()
    MsgBox "�i�ڊǗ��\�̃f�[�^���擾���܂�", vbInformation
    
    Application.ScreenUpdating = False
'    On Error GoTo ErrHdl
    Dim �i�ڋL���^�V�[�g  As String
    Dim �i�ڊǗ��\�Z�� As String
    Dim vPath As String
    �i�ڋL���^�V�[�g = �ݒ�擾("���i�ڋL���^", "�i�ڋL���^�i���W���[���j")
    �i�ڊǗ��\�Z�� = �ݒ�擾("���i�ڋL���^", "�i�ڊǗ��\�f�B���N�g���p�X")
    vPath = ThisWorkbook.Worksheets(�i�ڋL���^�V�[�g).Range(�i�ڊǗ��\�Z��).Value
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(vPath) Then
    
    Else
        MsgBox "�u�i�ڊǗ��\�v�t�@�C�������݂��܂���" _
            & vbCrLf & "�t�@�C���̃p�X���m�F���Ă�������", vbExclamation
        Exit Sub
    End If
    
    Dim WB�i�ڊǗ��\ As Workbook
    Dim WS�i�ڊǗ��\ As Worksheet
    Dim WS�R�[�h�\ As Worksheet
    
    Set WB�i�ڊǗ��\ = Workbooks.Open(Filename:=vPath, UpdateLinks:=0, ReadOnly:=True)
    Set WS�i�ڊǗ��\ = WB�i�ڊǗ��\.Worksheets("Sheet1")
    Set WS�R�[�h�\ = WB�i�ڊǗ��\.Worksheets(WSNAME_CODE)
    
    Dim �ŏI�s As Long
    Dim �ŏI�� As Long
    Dim �\�t��s As Long
    
    With WS�i�ڊǗ��\
        �ŏI�s = .Cells(.Rows.Count, 2).End(xlUp).Row
        �ŏI�� = .Cells(5, .Columns.Count).End(xlToLeft).Column
    End With
    Dim i As Long
    Dim WS�i�ڊǗ��V�[�g As Worksheet
    Dim WS�R�[�h���X�g As Worksheet
    
    Set WS�i�ڊǗ��V�[�g = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
    Set WS�R�[�h���X�g = ThisWorkbook.Worksheets(WSNAME_CODE)
    
    WS�i�ڊǗ��V�[�g.Range("A1").CurrentRegion.Offset(1).Clear
    WS�R�[�h���X�g.Cells.Clear
    WS�R�[�h�\.Cells.Copy Destination:=WS�R�[�h���X�g.Range("A1")
    
    �\�t��s = 2
    For i = 4 To �ŏI�s
        If WS�i�ڊǗ��\.Cells(i, 2).Value = "�i��" Then
            If WS�i�ڊǗ��\.Cells(i + 1, 2).Value <> "" Then
                With WS�i�ڊǗ��\
                    .Range(.Cells(i + 1, 2), .Cells(i + 2, �ŏI��)).Copy _
                        Destination:=WS�i�ڊǗ��V�[�g.Cells(�\�t��s, 1)
                End With
                �\�t��s = �\�t��s + 1
            End If
        End If
    Next
    
    For i = 2 To �\�t��s - 1
        With WS�i�ڊǗ��V�[�g
            .Cells(i, 2).NumberFormat = "@"
            .Cells(i, 2).Value = CStr(�R�[�h�ϊ�(.Cells(1, 2).Value, .Cells(i, 2).Value))
            .Cells(i, 3).NumberFormat = "@"
            .Cells(i, 3).Value = �R�[�h�ϊ�(.Cells(1, 3).Value, .Cells(i, 3).Value)
            .Cells(i, 4).NumberFormat = "@"
            .Cells(i, 4).Value = �R�[�h�ϊ�(.Cells(1, 4).Value, .Cells(i, 4).Value)
            .Cells(i, 11).NumberFormat = "@"
            .Cells(i, 11).Value = �R�[�h�ϊ�(.Cells(1, 11).Value, .Cells(i, 11).Value)
            .Cells(i, 12).NumberFormat = "@"
            .Cells(i, 12).Value = �R�[�h�ϊ�(.Cells(1, 12).Value, .Cells(i, 12).Value)
            .Cells(i, 13).NumberFormat = "@"
            .Cells(i, 13).Value = �R�[�h�ϊ�(.Cells(1, 13).Value, .Cells(i, 13).Value)
'            .Cells(i, 17).Value = �R�[�h�ϊ�(.Cells(1, 17).Value, .Cells(i, 17).Value)
'            .Cells(i, 18).Value = �R�[�h�ϊ�(.Cells(1, 18).Value, .Cells(i, 18).Value)
'            .Cells(i, 19).Value = �R�[�h�ϊ�(.Cells(1, 19).Value, .Cells(i, 19).Value)
            .Cells(i, 20).NumberFormat = "@"
            .Cells(i, 20).Value = �R�[�h�ϊ�(.Cells(1, 20).Value, .Cells(i, 20).Value)
            .Cells(i, 21).NumberFormat = "@"
            .Cells(i, 21).Value = �R�[�h�ϊ�(.Cells(1, 21).Value, .Cells(i, 21).Value)
            .Cells(i, 22).NumberFormat = "@"
            .Cells(i, 22).Value = �R�[�h�ϊ�(.Cells(1, 22).Value, .Cells(i, 22).Value)
            .Cells(i, 23).NumberFormat = "@"
            .Cells(i, 23).Value = �R�[�h�ϊ�(.Cells(1, 23).Value, .Cells(i, 23).Value)
            .Cells(i, 25).NumberFormat = "@"
            .Cells(i, 25).Value = �R�[�h�ϊ�(.Cells(1, 25).Value, .Cells(i, 25).Value)
            
        End With
    Next
    ���͋K���ݒ�
    ���͋K���폜
    
    MsgBox "�������I�����܂���", vbInformation
ExitHdl:
    On Error Resume Next
    WB�i�ڊǗ��\.Close False
    On Error GoTo 0
    
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
    Resume ExitHdl
End Sub

'���͋K���̐ݒ�
Private Sub ���͋K���ݒ�()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    
    Dim �ŏI�s As Long
    With ThisWorkbook.Worksheets(WSNAME_HINMOKU)
        �ŏI�s = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long, j As Long
    For i = 1 To sh.UsedRange.Rows.Count
        For j = 1 To sh.UsedRange.Columns.Count
            If sh.Cells(i, j).Value = "�i��" Then
            With sh.Cells(i + 1, j).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop _
                    , Operator:=xlEqual, Formula1:="=�i�ڊǗ��\!$A$2:$A$" & �ŏI�s
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = False
            End With
            End If
        Next
    Next
End Sub

'���͋K���̍폜
Private Sub ���͋K���폜()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
        
    Dim �ŏI�� As Long
    Dim �ŏI�s As Long
    With sh
        �ŏI�� = .Cells(1, .Columns.Count).End(xlToLeft).Column
        �ŏI�s = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    With sh.Range(sh.Cells(2, 1), sh.Cells(�ŏI�s, �ŏI��)).Validation
        .Delete
    End With
End Sub

Private Sub �R�[�h�ϊ�Test()
    Debug.Print �R�[�h�ϊ�("���B�����敪", "DC�N�_�H��")
End Sub
'�����̃R�[�h�ϊ��i�u�R�[�h�ꗗ�v�V�[�g�Q�Ɓj
Private Function �R�[�h�ϊ�(ByVal �Ώ� As String _
    , ByVal �l As Variant) As Variant
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_CODE)
    
    Dim �ΏۃZ�� As Range
    Set �ΏۃZ�� = sh.Rows(1).Find(�Ώ�)
    If �ΏۃZ�� Is Nothing Then
        �R�[�h�ϊ� = False
        Exit Function
    End If
    Dim �Ώ۔͈� As Range
    Set �Ώ۔͈� = �ΏۃZ��.CurrentRegion
    Dim i As Long
    For i = 2 To �Ώ۔͈�.Rows.Count
        If �Ώ۔͈�.Cells(i, 1).Value = �l Then
            �R�[�h�ϊ� = �Ώ۔͈�.Cells(i, 2).Value
        End If
    Next
End Function



