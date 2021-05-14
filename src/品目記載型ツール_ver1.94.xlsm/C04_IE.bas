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

Private Sub IE�O�ʃe�X�g()
    IE�O�� "���̃y�[�W"
End Sub
'Internet Explorer���őO�ʂɕ\������
Public Sub IE�O��(ByVal Title As Variant)
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

'�Ώۂ̃V�[�g���e�Ɋ�Â�IE�֘A�̏������s��
Public Sub IE����(ByVal �ΏۃV�[�g As Worksheet)
    Dim �ŏI�s As Long
    With �ΏۃV�[�g
        �ŏI�s = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 4 To �ŏI�s
        Select Case �ΏۃV�[�g.Cells(i, 4).Value
            Case "a"
                
            Case "input"
                ClickButton �ΏۃV�[�g.Cells(i, 6).Value
        End Select
    Next
End Sub

'�N������IE���n���h������
Public Function GetIEWindow(ByVal Title As Variant) As Variant
    Call ieWaitCheck
    
    Dim objIE As Object
    Set objIE = GetActiveIE(Title)
    
    If objIE Is Nothing Then
        GetIEWindow = False
    End If
    Set GetIEWindow = objIE
End Function

'URL���w�肵�ċN������IE�擾����
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
'�{�^���N���b�N�iinput�^�O�j
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

'�J���iA�^�O�j
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

'�f�[�^����
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

'�f�[�^�擾
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

'���ϔԍ��擾
Public Function ���ϔԍ��擾() As Variant
    Dim objIE As Object
    Call ieWaitCheck
    
    Set objIE = GetIEWindow("���Ϗ��Q��/SSIS")
'    Dim tb As Object
    Dim tr As Object
    Dim td As Object
    For Each tr In objIE.document.getElementsByTagName("tbody")(7).getElementsByTagName("tr")
        For Each td In tr.getElementsByTagName("td")
            If td.innerText Like "K*" Then
                ���ϔԍ��擾 = td.innerText
                Exit Function
            End If
        Next td
    Next tr
End Function

Private Sub IE�N���m�F�e�X�g()
    Debug.Print IE�N���m�F
End Sub
'IE���N�����Ă��邩�m�F����
Public Function IE�N���m�F() As Boolean
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
        IE�N���m�F = False
    Else
        IE�N���m�F = True
    End If
End Function

'IE���I������
Public Sub IE�I��()
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
