Attribute VB_Name = "����_IE����"
Option Explicit

Public Sub ieCheckSSISLogin(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)
    
    Dim �������[�h1 As String, �������[�h2 As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1

        Dim oTAG As Object

        Dim errorMsgFlg As Boolean: errorMsgFlg = False
        Dim ���b�Z�[�W As String

        For Each oTAG In oIE.document.getElementsByTagName("html")
            If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) <> 0 Then errorMsgFlg = True
            If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) <> 0 Then errorMsgFlg = True
            Err.Clear
            
            If errorMsgFlg = True Then
                ���b�Z�[�W = "���O�C���Ɏ��s���܂����B�I�[�g�p�C���b�g�͒�~���܂����B"
                MessageBox 0, ���b�Z�[�W, "�m�F", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION
                End
            End If
        Next
continue1:

    Next
    On Error GoTo 0


End Sub

'-------------------------
'IE���݃`�F�b�N
'-------------------------
Public Sub ieExistCheck()

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
        If oWin.Name = "Internet Explorer" Then
            MsgBox "InternetExplorer��S�����Ƃ��Ă���A�ēx�A���s���Ă��������B", vbExclamation
            End
        End If
    Next

End Sub

'-------------------------
'�A�N�V������̕\�������ҋ@
'-------------------------
Public Sub ieWaitCheck()
   
    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    
ReTry:
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin
            
        Dim timeOut As Date
        
        '���S�Ƀy�[�W���\�������܂őҋ@����
        timeOut = Now + TimeSerial(0, 0, 20)
         
'        Do While oIE.Busy = True Or (oIE.readyState < 4 And oIE.readyState > 0)
        Do While oIE.Busy = True Or (oIE.readyState < 4)
            If Err.Number <> 0 Then GoTo ReTry
'            DoEvents
            Sleep 250
            If Now > timeOut Then
                Debug.Print "�yA�zoIE.ReadyState:" & oIE.readyState & " hWnd:" & oWin.hWnd & " LocationName:" & oWin.LocationName
                oIE.Refresh
                timeOut = Now + TimeSerial(0, 0, 20)
            End If
        Loop
        
        Sleep 250
        timeOut = Now + TimeSerial(0, 0, 20)
        
        Do While oIE.document.readyState <> "complete"
            If Err.Number <> 0 Then GoTo ReTry
'            DoEvents
            Sleep 250
            If Now > timeOut Then
                Debug.Print "�yB�zoIE.ReadyState:" & oIE.readyState & " hWnd:" & oWin.hWnd & " LocationName:" & oWin.LocationName
                oIE.Refresh
                Sleep 1000
                timeOut = Now + TimeSerial(0, 0, 20)
            End If
        Loop
        On Error GoTo 0

continue1:
    Next

End Sub

'-------------------------
'�A�N�V������̕\�������ҋ@
'-------------------------
'Public Sub ieWaitCheck(ByVal oIE As Object)
'
'    Dim timeOut As Date
'
'    '���S�Ƀy�[�W���\�������܂őҋ@����
'    timeOut = Now + TimeSerial(0, 0, 20)
'
'    Debug.Print "oIE.ReadyState:" & oIE.readyState
'
'    Do While oIE.Busy = True Or (oIE.readyState < 4 And oIE.readyState > 0)
'        DoEvents
'        Sleep 250
'        If Now > timeOut Then
'            oIE.Refresh
'            timeOut = Now + TimeSerial(0, 0, 20)
'        End If
'    Loop
'
'    Sleep 250
'
'    timeOut = Now + TimeSerial(0, 0, 20)
'
'    Do While oIE.document.readyState <> "complete"
'        DoEvents
'        Sleep 1
'        If Now > timeOut Then
'            oIE.Refresh
'            timeOut = Now + TimeSerial(0, 0, 20)
'        End If
'    Loop
'
'End Sub
