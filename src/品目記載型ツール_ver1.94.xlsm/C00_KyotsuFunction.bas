Attribute VB_Name = "C00_KyotsuFunction"
Option Explicit
Option Private Module

'�K�{���ڊm�F
Public Function Fnc�K�{���ڊm�F(ByVal sh As Worksheet) As Boolean
    Dim temp As String
    If sh.Range("D3").Value = "" Then
        temp = temp & sh.Range("D2").Value & ";"
    End If

    If sh.Range("G3").Value = "" Then
        temp = temp & sh.Range("G2").Value & ";"
    End If
    If sh.Range("H3").Value = "" Then
        temp = temp & sh.Range("H2").Value & ";"
    End If
    If sh.Range("K3").Value = "" Then
        temp = temp & sh.Range("K2").Value & ";"
    End If
    If sh.Range("O3").Value = "" Then
        temp = temp & sh.Range("O2").Value & ";"
    End If
    
    If Len(temp) > 0 Then
        temp = Left(temp, Len(temp) - 1)
        MsgBox temp & "�͕K�{���ڂł��B�m�F���Ă�������", vbExclamation
        Fnc�K�{���ڊm�F = False
    Else
        Fnc�K�{���ڊm�F = True
    End If
    Exit Function
End Function

Private Sub �ϐ��擾Test()
    Debug.Print �ϐ��擾("�u���E�U�Ǘ��p", "A17", 3, "���C�����j���[")
End Sub
'�ϐ��擾�p
Public Function �ϐ��擾(ByVal �ΏۃV�[�g As String _
    , ByVal �Ώۃe�[�u�� As String _
    , ByVal �ΏۃJ���� As Long _
    , ByVal �L�[���[�h As String) As Variant
    
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(�ΏۃV�[�g)
    Dim Target As Range
    Set Target = sh.Range(�Ώۃe�[�u��).CurrentRegion
    
    For i = 1 To Target.Rows.Count
        If Target.Cells(i, 2).Value = �L�[���[�h Then
            �ϐ��擾 = Target.Cells(i, �ΏۃJ����).Value
            Exit Function
        End If
    Next
End Function

Private Sub �L���u���e�X�g()
    MsgBox �L���u��("����", "�y�����z�������_�@ACL�ǉ��Ή�", "�V���l")
End Sub
'�L����u������
Public Function �L���u��(ByVal �L�� As String _
    , ByVal �Ώە����� As String _
    , ByVal �u�������� As String) As String
    
    Dim temp As String
    temp = Replace(�Ώە�����, �L��, �u��������)
    �L���u�� = temp
End Function

'�Ώۂ̃u�b�N���擾
Public Function GetTargetBook(ByVal vPath As String) As Workbook
    Dim vBook As Workbook
    Dim temp As Workbook
    
    On Error GoTo ErrHdl
    For Each temp In Workbooks
        If temp.FullName = vPath Then
            Set GetTargetBook = temp
            Exit Function
        End If
    Next
    Set GetTargetBook = Workbooks.Open(Filename:=vPath)
ExitHdl:
    Exit Function
ErrHdl:
    Resume ExitHdl
End Function

'UsedRange�΍�
'UsedRange�́AA1���N�_�ɂȂ�Ƃ͌���Ȃ�����
Public Sub CorrectUsedRange(ByVal sh As Worksheet _
    , ByVal IsSetting As Boolean)
    If IsSetting Then
        If Trim(sh.Range("A1").Value) = vbNullString Then
            sh.Range("A1").Value = "Dummy"
        End If
    Else
        If sh.Range("A1").Value = "Dummy" Then
            sh.Range("A1").Value = vbNullString
        End If
    End If
End Sub


'�t�@�C���̑��݃`�F�b�N�e�X�g
Private Sub HasTargetFileTest()
    Dim vPath As String
    Dim vFileName As String
    
    vPath = ThisWorkbook.Path
    vFileName = "Test.csv"
    Debug.Print "TRUE:" & HasTargetFile(vPath, vFileName)
    vPath = vPath & "TEST"
    Debug.Print "FALSE:" & HasTargetFile(vPath, vFileName)
End Sub

'�t�@�C���̑��݃`�F�b�N
Public Function HasTargetFile(ByVal vPath As String _
    , ByVal vFileName As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim TargetPath As String
    TargetPath = vPath & "\" & vFileName
    If fso.FileExists(TargetPath) Then
        HasTargetFile = True
    Else
        HasTargetFile = False
    End If
End Function

'�t�H���_�̑��݃`�F�b�N�e�X�g
Private Sub HasTargetFolderTest()
    Dim vPath As String
    Dim vFileName As String
    
    vPath = ThisWorkbook.Path
    Debug.Print "TRUE:" & HasTargetFolder(vPath)
    vPath = vPath & "TEST"
    Debug.Print "FALSE:" & HasTargetFolder(vPath)
End Sub

'�t�H���_�̑��݃`�F�b�N
Public Function HasTargetFolder(ByVal vPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(vPath) Then
        HasTargetFolder = True
    Else
        HasTargetFolder = False
    End If
End Function

Private Sub GetFileCountTest()
    Debug.Print GetFileCount(ThisWorkbook.Path, "xlsx")
End Sub
'�w�肵�������̃t�@�C���̐����J�E���g
Public Function GetFileCount(ByVal vPath As String _
    , ByVal vExt As String) As Long
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim num As Long
    Dim temp As Object
    With fso
        For Each temp In .GetFolder(vPath).Files
            If UCase(.GetExtensionName(temp.Path)) _
                = UCase(vExt) Then
                num = num + 1
            End If
        Next
    End With
    GetFileCount = num
End Function


