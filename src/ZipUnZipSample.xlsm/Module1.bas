Attribute VB_Name = "Module1"
Option Explicit

'https://vbabeginner.net/zip-comp-decomp/
Sub UnZipTest()
    Dim sZipPath    As String
    Dim sExpandPath As String
    Dim bResult     As Boolean
    
    sZipPath = "C:\Users\21501173\Desktop\画像テスト.zip"
    sExpandPath = "C:\Users\21501173\Desktop"
    
    bResult = UnZip(sZipPath, sExpandPath)
    
    If bResult = True Then
        Call MsgBox("解凍しました。" & vbCr & sExpandPath, vbOKOnly, "解凍完了")
    Else
        Call MsgBox("解凍失敗しました。" & vbCr & sZipPath, vbOKOnly, "解凍失敗")
    End If
End Sub

Function UnZip(a_sZipPath As String, a_sExpandPath As String) As Boolean
    Dim sh      As New IWshRuntimeLibrary.WshShell
    Dim ex      As WshExec
    Dim sCmd    As String
    
    '// 半角スペースをバッククォートでエスケープ
    a_sZipPath = Replace(a_sZipPath, " ", "` ")
    a_sExpandPath = Replace(a_sExpandPath, " ", "` ")
    
    '// Expand-Archive：解凍コマンド
    '// -Path：フォルダパスまたはファイルパスを指定する。
    '// -DestinationPath：生成ファイルパスを指定する。
    '// -Force：生成ファイルが既に存在している場合は上書きする
    sCmd = "Expand-Archive -Path " & a_sZipPath & " -DestinationPath " & a_sExpandPath & " -Force"
    
    '// コマンド実行
    Set ex = sh.Exec("powershell -NoLogo -ExecutionPolicy RemoteSigned -Command " & sCmd)
    
    '// コマンド失敗時
    If ex.Status = WshFailed Then
        '// 戻り値に異常を返す
        UnZip = False
        
        '// 処理を抜ける
        Exit Function
    End If
    
    '// コマンド実行中は待ち
    Do While ex.Status = WshRunning
        DoEvents
    Loop
    
    '// 戻り値に正常を返す
    UnZip = True
End Function


'https://www.ka-net.org/blog/?p=7605
Public Sub ZipSample()
  ZipFileOrFolder "C:\Test\Files" 'フォルダ圧縮
  MsgBox "処理が終了しました。", vbInformation + vbSystemModal
End Sub
 
Public Sub UnZipSample()
  UnZipFile "C:\Users\21501173\Desktop\画像テスト.zip"
  MsgBox "処理が終了しました。", vbInformation + vbSystemModal
End Sub
 
Public Sub ZipFileOrFolder(ByVal SrcPath As Variant, _
                           Optional ByVal DestFolderPath As Variant = "")
'ファイル・フォルダをZIP形式で圧縮
'SrcPath：元ファイル・フォルダ
'DestFolderPath：出力先、指定しない場合は元ファイル・フォルダと同じ場所
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
    '空のZIPファイル作成
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
'ZIPファイルを解凍
'SrcPath：元ファイル
'DestFolderPath：出力先、指定しない場合は元ファイルと同じ場所
'※出力先に同名ファイルがあった場合はユーザー判断で処理
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
