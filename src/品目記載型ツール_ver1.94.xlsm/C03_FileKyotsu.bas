Attribute VB_Name = "C03_FileKyotsu"
Option Explicit
Option Private Module


'Excel�t�@�C���̃t���p�X����e�L�X�g�t�@�C���̃p�X���쐬
Public Function GetTextFileFullPath(ByVal vPath As String) As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim vFileName As String
    Dim TempPath As String
    TempPath = fso.GetParentFolderName(vPath)
    vFileName = fso.GetBaseName(vPath)
    GetTextFileFullPath = TempPath & "\" & vFileName & ".txt"
End Function

'�e�L�X�g�t�@�C���̓��e���擾
Public Function TextFileData(ByVal vPath As String) As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim buf As String
    With fso
        With .GetFile(vPath).OpenAsTextStream
            buf = .ReadAll
            .Close
        End With
    End With
    TextFileData = buf
End Function


Private Sub �Ώۃt�@�C���ړ��e�X�g()
    �Ώۃt�@�C���ړ� "C:\Users\21501173\Desktop\NESIC�l\Data\�yXXX�z���_�@ACL�ǉ��Ή�.csv"
End Sub
'�t�@�C���̈ړ�
Public Sub �Ώۃt�@�C���ړ�(ByVal vPath As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim vFileName As String
    Dim TempPath As String
    TempPath = fso.GetParentFolderName(vPath)
    vFileName = fso.GetBaseName(vPath)
    
    Dim TargetPath As String
    TargetPath = TempPath & "\old"
    If fso.FolderExists(TargetPath) = False Then
        fso.CreateFolder TargetPath
    End If
    fso.MoveFile vPath, TargetPath & "\" & vFileName & "_" & Format(Now, "YYYYMMDD_HHNN") & ".csv"
    fso.MoveFile TempPath & "\" & vFileName & ".txt", TargetPath & "\" & vFileName & "_" & Format(Now, "YYYYMMDD_HHNN") & ".txt"
End Sub

Private Sub ���ϔԍ��t�@�C���ۑ��e�X�g()
    ���ϔԍ��t�@�C���ۑ� "C:\Users\21501173\Desktop\NESIC�l\Data\old" _
        , "����", "K00001111-111"
End Sub

'���ϔԍ��t�@�C����ۑ�
Public Sub ���ϔԍ��t�@�C���ۑ�(ByVal vPath As String _
    , ByVal vSubject As String, ByVal Knumber As String)
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    If fso.FolderExists(vPath) = False Then
        fso.CreateFolder vPath
    End If
    Open vPath & "\" & vSubject & "_" & Knumber & ".txt" For Output As #1

    Print #1, Knumber
 
    Close #1
End Sub

