Attribute VB_Name = "C03_FileKyotsu"
Option Explicit
Option Private Module


'Excelファイルのフルパスからテキストファイルのパスを作成
Public Function GetTextFileFullPath(ByVal vPath As String) As String
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim vFileName As String
    Dim TempPath As String
    TempPath = fso.GetParentFolderName(vPath)
    vFileName = fso.GetBaseName(vPath)
    GetTextFileFullPath = TempPath & "\" & vFileName & ".txt"
End Function

'テキストファイルの内容を取得
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


Private Sub 対象ファイル移動テスト()
    対象ファイル移動 "C:\Users\21501173\Desktop\NESIC様\Data\【XXX】拠点　ACL追加対応.csv"
End Sub
'ファイルの移動
Public Sub 対象ファイル移動(ByVal vPath As String)
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

Private Sub 見積番号ファイル保存テスト()
    見積番号ファイル保存 "C:\Users\21501173\Desktop\NESIC様\Data\old" _
        , "件名", "K00001111-111"
End Sub

'見積番号ファイルを保存
Public Sub 見積番号ファイル保存(ByVal vPath As String _
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

