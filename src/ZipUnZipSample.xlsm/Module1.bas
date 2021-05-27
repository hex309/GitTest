Attribute VB_Name = "Module1"
Option Explicit

'https://vbabeginner.net/zip-comp-decomp/
Sub UnZipTest()
    Dim sZipPath    As String
    Dim sExpandPath As String
    Dim bResult     As Boolean
    
    sZipPath = "C:\Users\21501173\Desktop\�摜�e�X�g.zip"
    sExpandPath = "C:\Users\21501173\Desktop"
    
    bResult = UnZip(sZipPath, sExpandPath)
    
    If bResult = True Then
        Call MsgBox("�𓀂��܂����B" & vbCr & sExpandPath, vbOKOnly, "�𓀊���")
    Else
        Call MsgBox("�𓀎��s���܂����B" & vbCr & sZipPath, vbOKOnly, "�𓀎��s")
    End If
End Sub

Function UnZip(a_sZipPath As String, a_sExpandPath As String) As Boolean
    Dim sh      As New IWshRuntimeLibrary.WshShell
    Dim ex      As WshExec
    Dim sCmd    As String
    
    '// ���p�X�y�[�X���o�b�N�N�H�[�g�ŃG�X�P�[�v
    a_sZipPath = Replace(a_sZipPath, " ", "` ")
    a_sExpandPath = Replace(a_sExpandPath, " ", "` ")
    
    '// Expand-Archive�F�𓀃R�}���h
    '// -Path�F�t�H���_�p�X�܂��̓t�@�C���p�X���w�肷��B
    '// -DestinationPath�F�����t�@�C���p�X���w�肷��B
    '// -Force�F�����t�@�C�������ɑ��݂��Ă���ꍇ�͏㏑������
    sCmd = "Expand-Archive -Path " & a_sZipPath & " -DestinationPath " & a_sExpandPath & " -Force"
    
    '// �R�}���h���s
    Set ex = sh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & sCmd)
    
    '// �R�}���h���s��
    If ex.Status = WshFailed Then
        '// �߂�l�Ɉُ��Ԃ�
        UnZip = False
        
        '// �����𔲂���
        Exit Function
    End If
    
    '// �R�}���h���s���͑҂�
    Do While ex.Status = WshRunning
        DoEvents
    Loop
    
    '// �߂�l�ɐ����Ԃ�
    UnZip = True
End Function


'https://www.ka-net.org/blog/?p=7605
Public Sub ZipSample()
  ZipFileOrFolder "C:\Test\Files" '�t�H���_���k
  MsgBox "�������I�����܂����B", vbInformation + vbSystemModal
End Sub
 
Public Sub UnZipSample()
  UnZipFile "C:\Users\21501173\Desktop\�摜�e�X�g.zip"
  MsgBox "�������I�����܂����B", vbInformation + vbSystemModal
End Sub
 
Public Sub ZipFileOrFolder(ByVal SrcPath As Variant, _
                           Optional ByVal DestFolderPath As Variant = "")
'�t�@�C���E�t�H���_��ZIP�`���ň��k
'SrcPath�F���t�@�C���E�t�H���_
'DestFolderPath�F�o�͐�A�w�肵�Ȃ��ꍇ�͌��t�@�C���E�t�H���_�Ɠ����ꏊ
  Dim DestFilePath As Variant
   
  With CreateObject("Scripting.FileSystemObject")
    If IsFolder(DestFolderPath) = False Then
      If IsFolder(SrcPath) = True Then
        DestFolderPath = SrcPath
      ElseIf IsFile(SrcPath) = True Then
        DestFolderPath = .GetFile(SrcPath).ParentFolder.Path
      Else: Exit Sub
      End If
    End If
    DestFilePath = AddPathSeparator(DestFolderPath) & _
                     .GetBaseName(SrcPath) & ".zip"
    '���ZIP�t�@�C���쐬
    With .CreateTextFile(DestFilePath, True)
      .Write ChrW(&H50) & ChrW(&H4B) & ChrW(&H5) & ChrW(&H6) & String(18, ChrW(0))
      .Close
    End With
  End With
   
  With CreateObject("Shell.Application")
    With .Namespace(DestFilePath)
      .CopyHere SrcPath
      While .Items.Count < 1
        DoEvents
      Wend
    End With
  End With
End Sub
 
Public Sub UnZipFile(ByVal SrcPath As Variant, _
                     Optional ByVal DestFolderPath As Variant = "")
'ZIP�t�@�C������
'SrcPath�F���t�@�C��
'DestFolderPath�F�o�͐�A�w�肵�Ȃ��ꍇ�͌��t�@�C���Ɠ����ꏊ
'���o�͐�ɓ����t�@�C�����������ꍇ�̓��[�U�[���f�ŏ���
  With CreateObject("Scripting.FileSystemObject")
    If .FileExists(SrcPath) = False Then Exit Sub
    If LCase(.GetExtensionName(SrcPath)) <> "zip" Then Exit Sub
    If IsFolder(DestFolderPath) = False Then
      DestFolderPath = .GetFile(SrcPath).ParentFolder.Path
    End If
  End With
   
  With CreateObject("Shell.Application")
    .Namespace(DestFolderPath).CopyHere .Namespace(SrcPath).Items
  End With
End Sub
 
Private Function IsFolder(ByVal SrcPath As String) As Boolean
  IsFolder = CreateObject("Scripting.FileSystemObject").FolderExists(SrcPath)
End Function
 
Private Function IsFile(ByVal SrcPath As String) As Boolean
  IsFile = CreateObject("Scripting.FileSystemObject").FileExists(SrcPath)
End Function
 
Private Function AddPathSeparator(ByVal SrcPath As String) As String
  If Right(SrcPath, 1) <> ChrW(92) Then SrcPath = SrcPath & ChrW(92)
  AddPathSeparator = SrcPath
End Function
