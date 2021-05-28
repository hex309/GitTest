Attribute VB_Name = "Tester"
Option Explicit
Option Private Module

Private Sub resetTool()
    Initializer
    prepTest
End Sub

Private Sub Initializer()
    Dim sh As Worksheet
    
    Application.ScreenUpdating = False

    On Error Resume Next
    getCurrentRegion(LogSh.Cells(2, 1), 1, False).ClearContents
    getCurrentRegion(OldLogSh.Cells(2, 1), 1, False).ClearContents
    getCurrentRegion(SeminarSh.Cells(3, 1), 1, False).ClearContents
    getCurrentRegion(AccountSh.Cells(3, 1), 2, False).ClearContents
    getCurrentRegion(MailSettingSh.Cells(3, 1), 2, False).Resize(, 2).ClearContents
    ScenarioSh.Range("N2").ClearContents
    ScenarioSh.Range("N4").Value = True
    ScenarioSh.Range("N5").Value = True
    ScenarioSh.Range("N7").Valuxe = True
    ScenarioSh.Range("N8").Value = True
    ScenarioSh.Columns("L").ClearContents
    ScenarioSh.Range("C3:F12").ClearContents
    ScenarioSh.Range("I3:K12").ClearContents
    SettingSh.Range("B11:C20").ClearContents
    SettingSh.Range("B3").Value = "Downloads"
    MailSettingSh.Range("C3:C21").Value = False
    MailSettingSh.Range("D3:M21").Value = True
    
    On Error GoTo 0
    
    For Each sh In ThisWorkbook.Worksheets
        With sh
            .Activate
            .Cells(1, 1).Select
            ActiveWindow.ScrollRow = 1
            ActiveWindow.ScrollColumn = 1
        End With
    Next
    
    ScenarioSh.Activate
    
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub prepTest()
    Dim prbk As Workbook
    Dim tgtVal As Variant
    
    Application.ScreenUpdating = False
    
    Set prbk = Workbooks("テストパラメータ.xlsx")
    
    tgtVal = getCurrentRegion(prbk.Worksheets(1).Cells(1, 1)).Value
    
    With SeminarSh
        .Range(.Cells(3, 1), .Cells(UBound(tgtVal, 1) + 2, UBound(tgtVal, 2))).Value = tgtVal
    End With
    
    tgtVal = getCurrentRegion(prbk.Worksheets(2).Cells(1, 1)).Value
    
    With AccountSh
        .Range(.Cells(3, 1), .Cells(UBound(tgtVal, 1) + 2, UBound(tgtVal, 2))).Value = tgtVal
    End With
    
    SettingSh.Range("B3").Value = "../Downloads"
    
    MailSettingSh.Activate
    MailSettingSh.Range("A3:B3").Value = prbk.Worksheets(3).Range("A1:B1").Value
    
    ScenarioSh.Activate
    ScenarioSh.Range("N2").Value = "橋本　啓"
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub domTester()
    Dim objIE As InternetExplorer
    Dim docIE As HTMLDocument
    Dim objShell As Object
    Dim objWin As Object
    Dim address As String
    Dim i As Long
    
    Set opeLog = New Collection
    
    address = "reserveimport/confirm"
  
'#####Shellオブジェクトを作成する
    Set objShell = CreateObject("Shell.Application")
    For Each objWin In objShell.Windows
    
        If objWin.name = "Internet Explorer" Then
            If InStr(objWin.LocationURL, address) > 0 Then
                'InternetExplorerオブジェクトをセット
                Set objIE = objWin
                Exit For
            End If
        End If
    Next
    
    Set docIE = objIE.Document

'####テストしたい処理

    docIE.getElementById("comMiwsTopLink").click
    
normalFin:
'####Errorメッセージ
    
    For i = 1 To opeLog.Count
        Debug.Print opeLog(i)
    Next
    
    Set opeLog = Nothing
    
'    Stop
    
End Sub

Private Sub funcTester()
    Dim tgtSite As CorpSite
    Dim tgtSiteName As String
    Dim tgtCorpName As String
    Dim srcFilePath As String
    Dim i As Long

    tgtCorpName = "シンカ"
    tgtSiteName = "リクナビ"
    
    Set opeLog = New Collection
    Set mailInfo = New Collection
    
    Set tgtSite = New CorpSite
    
'####対象サイト用のIEラッパーを起動
    If Not tgtSite.setCorp(tgtCorpName, tgtSiteName) Then
        MsgBox "ieWrapper couldnot start!"
        GoTo errormsg
    End If
   
'####テストしたい処理
    loginRikuNavi tgtSite
    searchRikuNaviDt tgtSite, dataType.Seminar

'    loginMyNavi tgtSite
'    moveSearchWindow tgtSite, True
'    searchMyNaviDt tgtSite, dataType.Seminar
    Stop

errormsg:
'####Errorメッセージ
    For i = 1 To opeLog.Count
        Debug.Print opeLog(i)
    Next

    Set opeLog = Nothing
    
    Unload AlertBox
    
    If Not tgtSite Is Nothing Then tgtSite.quitAll
    
End Sub

Private Sub VBSTestTest()
    Debug.Print VBSTest("C:\Users\21501173\Desktop\")
End Sub

Private Function VBSTest(ByVal vPath As String) As Boolean
    Dim vbCode As String
    Dim fso As Object
    Dim txtStrm As Object
    Dim hWindow As LongPtr
    Dim path As String
    Dim timeOut As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    path = fso.BuildPath(vPath, "OpenDialog.vbs")
    
    vbCode = "MsgBox ""OK"""
            
    If Not fso.FileExists(path) Then
        Set txtStrm = fso.CreateTextFile(path)
        txtStrm.write vbCode
        txtStrm.Close
    End If
    
    Shell "WScript.exe " & """" & path & """"  '指定先ファイルを開く
    timeOut = Now + TimeValue("00:00:05")
    
    Do While hWindow = 0
        DoEvents
        'Application.Wait [now() + "00:00:01"]
    
        If Now > timeOut Then
'            opeLog.Add "アップロードするファイルの選択ダイアログを表示できませんでした。"
            Exit Function
        End If
        
    Loop
    
    On Error Resume Next
    fso.DeleteFile path
    On Error GoTo 0
    
    VBSTest = True
        
End Function
 
