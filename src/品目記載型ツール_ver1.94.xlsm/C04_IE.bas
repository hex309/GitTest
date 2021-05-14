Attribute VB_Name = "C04_IE"
Option Explicit
Option Private Module

#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
    Private Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function FindWindow Lib "User32.dll" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
    (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Long) As Long
#Else
    Private Declare Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
    Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
                             (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare Function FindWindow Lib "User32.dll" Alias "FindWindowA" _
                             (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryExA" _
                             (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Long) As Long
    
#End If
#If Win64 Then
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#Else
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const WM_COMMAND As Long = &H111&

Private Sub IE前面テスト()
    IE前面 "このページ"
End Sub
'Internet Explorerを最前面に表示する
Public Sub IE前面(ByVal Title As Variant)
    Call ieWaitCheck
    
    Dim objIE As Object
'    Set objIE = GetActiveIE(Title)
    Dim oShell As Object
    Dim oWin As Object
    
    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
        If oWin.Name = "Internet Explorer" Then
            Set objIE = oWin
            Exit For
        End If
    Next
    SetForegroundWindow objIE.hWnd
End Sub

'対象のシート内容に基づきIE関連の処理を行う
Public Sub IE処理(ByVal 対象シート As Worksheet)
    Dim 最終行 As Long
    With 対象シート
        最終行 = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 4 To 最終行
        Select Case 対象シート.Cells(i, 4).Value
            Case "a"
                
            Case "input"
                ClickButton 対象シート.Cells(i, 6).Value
        End Select
    Next
End Sub

'起動中のIEをハンドルする
Public Function GetIEWindow(ByVal Title As Variant) As Variant
    Call ieWaitCheck
    
    Dim objIE As Object
    Set objIE = GetActiveIE(Title)
    
    If objIE Is Nothing Then
        GetIEWindow = False
    End If
    Set GetIEWindow = objIE
End Function

'URLを指定して起動中のIE取得する
Public Function GetActiveIE(ByVal URL As String) As Object
    Dim objIE As Object
    Dim o As Object
    Call ieWaitCheck
    
    For Each o In GetObject("new:{9BA05972-F6A8-11CF-A442-00A0C90A8F39}") 'ShellWindows
        If LCase(TypeName(o)) = "iwebbrowser2" Then
            If LCase(TypeName(o.document)) = "htmldocument" Then
                If o.document.Title Like "*" & URL & "*" Then
                    Set GetActiveIE = o
                    Exit For
                End If
            End If
        End If
    Next
End Function

'Public Sub RefreshIE(ByVal objIE As Object)
'    objIE.Visible = False
'    Sleep 1000
'    DoEvents
'    objIE.Visible = True
'End Sub
'ボタンクリック（inputタグ）
Public Function ClickButton(ByVal ButtonName As String) As Boolean
    On Error GoTo ErrHdl
    Dim objIE As Object
    Set objIE = GetIEWindow("")
    
    Dim temp As Object
    Dim d
    With objIE.document.Frames("Main").document
        For Each temp In .getElementsByTagName("input")
            d = temp.Value
            If d = ButtonName Then
                temp.Click
                ClickButton = True
                Exit Function
            End If
        Next
    End With
ErrHdl:
    ClickButton = False
End Function

'開く（Aタグ）
Public Function ClickATag() As Object
    Dim objIE As Object
    Set objIE = GetIEWindow("")
    
    Dim temp As Object
    Dim d As Variant
    
    With objIE.document
'        .Script.setTimeout "javascript:return doFolderClick('item2','img2')"
'        .Script.setTimeout "javascript:return doFolderClick('item4','img4')"
        For Each temp In .getElementsByTagName("a")
            d = temp.innerText
            If d = "TEST" Then temp.Click
        Next
    End With
End Function

'データ入力
Public Function SetElementByID(ByVal ElementName As Variant _
    , ByVal vData As Variant) As Variant
    On Error GoTo ErrHdl
    Dim objIE As Object
    Set objIE = GetIEWindow("")
    
    With objIE.document
        .GetElementByID(ElementName).Value = vData
    End With
    SetElementByID = True
    Exit Function
ErrHdl:
    SetElementByID = False
End Function

'データ取得
Public Function GetElementByID(ByVal TargetWindow As String _
    , ByVal ElementName As Variant) As Variant
    On Error GoTo ErrHdl
    Dim objIE As Object
    Set objIE = GetIEWindow(TargetWindow)
    
    With objIE.document
        GetElementByID = .GetElementByID(ElementName).Value
    End With

    Exit Function
ErrHdl:
    GetElementByID = False
End Function

'見積番号取得
Public Function 見積番号取得() As Variant
    Dim objIE As Object
    Call ieWaitCheck
    
    Set objIE = GetIEWindow("見積情報参照/SSIS")
'    Dim tb As Object
    Dim tr As Object
    Dim td As Object
    For Each tr In objIE.document.getElementsByTagName("tbody")(7).getElementsByTagName("tr")
        For Each td In tr.getElementsByTagName("td")
            If td.innerText Like "K*" Then
                見積番号取得 = td.innerText
                Exit Function
            End If
        Next td
    Next tr
End Function

Private Sub IE起動確認テスト()
    Debug.Print IE起動確認
End Sub
'IEが起動しているか確認する
Public Function IE起動確認() As Boolean
    Dim oIE As Object
    Dim oShell As Object
    Dim oWin As Object
    
    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
        If oWin.Name = "Internet Explorer" Then
            Set oIE = oWin
            Exit For
        End If
    Next
    If oIE Is Nothing Then
        IE起動確認 = False
    Else
        IE起動確認 = True
    End If
End Function

'IEを終了する
Public Sub IE終了()
    Dim oIE As Object
    Dim oShell As Object
    Dim oWin As Object
    
    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
        If oWin.Name = "Internet Explorer" Then
            oWin.Quit
        End If
    Next
End Sub
