Attribute VB_Name = "ieOpe"
Option Explicit
Option Private Module

Private Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" _
        (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" _
        (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare PtrSafe Function SetForegroundWindow Lib "USER32" (ByVal hWnd As Long) As Long

Private Declare PtrSafe Function SendMessage Lib "user32.dll" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
        
Private Declare PtrSafe Function GetWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Sub setDialogByVBSTest()
    setDialogByVBS
End Sub
Public Function setDialogByVBS() As Boolean
    Dim vbCode As String
    Dim fso As Object
    Dim txtStrm As Object
    Dim hWindow As LongPtr
    Dim path As String
    Dim timeOut As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    path = fso.BuildPath(ThisWorkbook.path, "OpenDialog.vbs")
    
    vbCode = "Dim v " & vbCrLf & _
             "Dim objIE " & vbCrLf & _
             "Dim objShell" & vbCrLf & _
             "Dim objWin" & vbCrLf & _
             "" & vbCrLf & _
             "'Shell�I�u�W�F�N�g���쐬����" & vbCrLf & _
             "Set objShell = CreateObject(""Shell.Application"")" & vbCrLf & _
             "For Each objWin In objShell.Windows" & vbCrLf & _
             "" & vbCrLf & _
             "    If objWin.Name = ""Internet Explorer"" Then" & vbCrLf & _
             "        If instr(objWin.LocationURL,""wdi/index"")>0 then" & vbCrLf & _
             "            'InternetExplorer�I�u�W�F�N�g���Z�b�g" & vbCrLf & _
             "            Set objIE = objWin" & vbCrLf & _
             "            Exit For" & vbCrLf & _
             "        End If" & vbCrLf & _
             "    End If" & vbCrLf & _
             "Next" & vbCrLf & _
            "" & vbCrLf & _
            "For Each v In objIE.document.getElementsByName(""wdifile"")" & vbCrLf & _
            "    v.Click" & vbCrLf & _
            "Next"
            
    If Not fso.FileExists(path) Then
        Set txtStrm = fso.CreateTextFile(path)
        txtStrm.write vbCode
        txtStrm.Close
    End If
    
    Shell "WScript.exe " & """" & path & """"   '�w���t�@�C�����J��
    timeOut = Now + TimeValue("00:00:05")
    
    Do While hWindow = 0
        DoEvents
        'Application.Wait [now() + "00:00:01"]
    
        On Error Resume Next
        hWindow = FindWindow("#32770", "�A�b�v���[�h����t�@�C���̑I��")
        On Error GoTo 0
        
        If Now > timeOut Then
            opeLog.Add "�A�b�v���[�h����t�@�C���̑I���_�C�A���O��\���ł��܂���ł����B"
            Exit Function
        End If
        
        If cancelFlg Then
             opeLog.Add "�L�����Z������܂����B"
            Exit Function
        End If
    Loop
    
    On Error Resume Next
    fso.DeleteFile path
    On Error GoTo 0
    
    setDialogByVBS = True
        
End Function
 

Public Sub setFileName(ByVal fileName As String)
    Dim hWindow As LongPtr
    Dim timeOut As Date
    
    timeOut = Now + TimeSerial(0, 0, 30)
    
    Do
        hWindow = FindWindow("#32770", "�A�b�v���[�h����t�@�C���̑I��")
        
        If hWindow <> 0 Then setFileNameEx fileName, hWindow
        
        DoEvents
        Sleep 100
        
        If Now > timeOut Then
            Exit Do
        End If
        
    Loop Until hWindow = 0
    
End Sub

Private Sub setFileNameEx(ByVal fileName As String, ByVal hWindow As LongPtr)
    Dim hInputBox As LongPtr
    Dim hButton As LongPtr
    Dim hAlButton As LongPtr

    hInputBox = FindWindowEx(CLng(hWindow), 0&, "ComboBoxEx32", "")
    hInputBox = FindWindowEx(CLng(hInputBox), 0&, "ComboBox", "")
    hInputBox = FindWindowEx(CLng(hInputBox), 0&, "Edit", "")
    
    SendMessage CLng(hInputBox), &HC, 0, fileName
            
    hButton = FindWindowEx(CLng(hWindow), 0&, "Button", "�J��(&O)")
    
    '�{�^������'
    SendMessage CLng(hButton), &H6, 1, 0&  '�{�^�����A�N�e�B�u�ɂ���
    SendMessage CLng(hButton), &HF5, 0, 0& '�{�^�����N���b�N����
    
    hInputBox = FindWindowEx(CLng(hWindow), 0&, "DirectUIHWND", vbNullString)
    hInputBox = FindWindowEx(CLng(hInputBox), 0&, "CtrlNotifySink", vbNullString)
    
    Do Until hInputBox = 0
        hAlButton = FindWindowEx(CLng(hInputBox), 0&, "Button", "OK")
        
        If hAlButton = 0 Then
            hInputBox = GetWindow(CLng(hInputBox), 2)
        Else
            SendMessage CLng(hAlButton), &H6, 1, 0&
            SendMessage CLng(hAlButton), &HF5, 0, 0&
            Exit Do
        End If
    Loop

End Sub



'// ------------------------------------------------------------------------------------------------------------------------
'//  �v���V�[�W�����@�FIePost
'//  �@�\�@�@�@�@�@�@�F�w�肳�ꂽURL�Ƀf�[�^��POST����
'//  �����@�@�@�@�@�@�FobjIE�FInternetExplorer�I�u�W�F�N�g���w��
'//                  �FTargetURL�FPOST������URL�̕�������w��
'//                  �FViewFlg�F�ȗ��\�A�uTrue�v���K��l
'//                  �F�uTrue�v����IE�\���A�uFalse�v����IE��\��
'//  �߂�l�@�@�@�@�@�Fboolean
'//  �쐬�ҁ@�@�@�@�@�FShingo Maekawa
'//  �쐬���@�@�@�@�@�F2017/8/2
'//  ���l�@�@�@�@�@�@�F�T�u���[�`����
'//  �X�V���F���e�@�@�Fhttps://support.microsoft.com/ja-jp/help/174923/how-to-use-the-postdata-parameter-in-webbrowser-control
'// ------------------------------------------------------------------------------------------------------------------------

Public Function IePost(ByRef objIE As Object, _
                   ByVal TargetURL As String, _
                   Optional ByVal PostData As String = vbNullString, _
                   Optional ByVal ViewFlg As Boolean = True, _
                   Optional ByVal Headers As String = "Content-Type: application/x-www-form-urlencoded" & vbCrLf, _
                   Optional ByVal Flags As Long = 0, _
                   Optional ByVal TargetFrame As String = vbNullString) As Boolean
                   
    Dim bPostData() As Byte
    
    'IE(InternetExplorer)�I�u�W�F�N�g�������ꍇ�쐬����
    If objIE Is Nothing Then
        On Error Resume Next
        Set objIE = CreateObject("InternetExplorer.Application")
        On Error GoTo 0
    End If
    
    If objIE Is Nothing Then
        opeLog.Add "�C���^�[�l�b�g�G�N�X�v���[�����N���ł��܂���ł����B"
        Exit Function
    End If
    
    'IE(InternetExplorer)��\���E��\��
    objIE.visible = ViewFlg
    
    'PostData���o�C�i����
    bPostData = StrConv(PostData, vbFromUnicode)
    
    '�w�肵��URL�̃y�[�W��\������
    objIE.navigate TargetURL, Flags, TargetFrame, bPostData, Headers
 
    'IE(InternetExplorer)�����S�\�������܂őҋ@
    IECheck objIE
    
    IePost = True
    
End Function


'// ------------------------------------------------------------------------------------------------------------------------
'//  �v���V�[�W�����@�FDLReadyCheck
'//  �@�\�@�@�@�@�@�@�F�_�E�����[�h�ҋ@��Ԃ��`�F�b�N�iBusyState=False,Document.ReadyState=complete�j
'//  �����@�@�@�@�@�@�FobjIE�FInternetExplorer�I�u�W�F�N�g���w��
'//  �߂�l�@�@�@�@�@�F�Ȃ�
'//  �쐬�ҁ@�@�@�@�@�FShingo Maekawa
'//  �쐬���@�@�@�@�@�F2017/08/01
'//  ���l�@�@�@�@�@�@�F
'//  �X�V���F���e�@�@�F
'// ------------------------------------------------------------------------------------------------------------------------

Public Sub DLReadyCheck(ByRef objIE As Object)
    Dim timeOut As Date

    '���S�Ƀy�[�W���\�������܂őҋ@����
    timeOut = Now + TimeSerial(0, 3, 0)

    Do While objIE.Busy = True Or objIE.ReadyState < 3
        DoEvents
        Application.Wait [now() + "00:00:00.5"]
        
        If Now > timeOut Then
            MsgBox "TimeOut@DLWait"
        End If
    Loop

    timeOut = Now + TimeSerial(0, 0, 20)
    
    On Error Resume Next
    Do While objIE.Document.ReadyState <> "complete"
        DoEvents
        Application.Wait [now() + "00:00:00.5"]
        
        If Now > timeOut Then
            MsgBox "TimeOut@DLWait"
        End If
    Loop
    On Error GoTo 0
    
End Sub

Public Function IEPageMoveCheck(ByRef objIE As Object, ByVal chkUrl As String) As Boolean
    Dim timeOut As Date

    '�y�[�W�̑J�ڊ����܂őҋ@����
    timeOut = Now + TimeSerial(0, 3, 0)
    
    Do While True
        Application.Wait [now() + "00:00:00.5"]
        
        If Now > timeOut Then
            opeLog.Add "���͑҂��Ń^�C���A�E�g���܂����B"
            Exit Do
        End If
    
        IECheck objIE
    
        If InStr(objIE.LocationURL, chkUrl) > 0 Then
            Exit Do
        End If
    Loop
    
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  �v���V�[�W�����@�FIECheck
'//  �@�\�@�@�@�@�@�@�FWeb�y�[�W�����S�ɓǂݍ��܂��܂őҋ@����
'//  �����@�@�@�@�@�@�FobjIE�FInternetExplorer�I�u�W�F�N�g���w��
'//  �߂�l�@�@�@�@�@�F�Ȃ�
'//  �쐬�ҁ@�@�@�@�@�FShingo Maekawa
'//  �쐬���@�@�@�@�@�F2017/08/01
'//  ���l�@�@�@�@�@�@�F
'//  �X�V���F���e�@�@�F
'// ------------------------------------------------------------------------------------------------------------------------

Public Sub IECheck(ByRef objIE As Object)
    Dim timeOut As Date

    '���S�Ƀy�[�W���\�������܂őҋ@����
    timeOut = Now + TimeSerial(0, 0, 20)

    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        Application.Wait [now() + "00:00:00.5"]
        
        If Now > timeOut Then
            'objIE.Refresh
            timeOut = Now + TimeSerial(0, 0, 20)
        End If
    Loop

    timeOut = Now + TimeSerial(0, 0, 20)
    
    On Error Resume Next
    Do While objIE.Document.ReadyState <> "complete"
        DoEvents
        Application.Wait [now() + "00:00:00.5"]
        
        If Now > timeOut Then
            'objIE.Refresh
            timeOut = Now + TimeSerial(0, 0, 20)
        End If
    Loop
    On Error GoTo 0
    
End Sub

Public Function pushSaveButton(ByRef objIE As Object) As Boolean
    Dim AutomationObj As IUIAutomation
    Dim WindowElement As IUIAutomationElement
    Dim saveButton As IUIAutomationElement
    Dim hWnd As LongPtr
    
    Set AutomationObj = New CUIAutomation
    
    hWnd = objIE.hWnd
    hWnd = FindWindowEx(CLng(hWnd), 0, "Frame Notification Bar", vbNullString)
    If hWnd = 0 Then Exit Function
    
    'SetForegroundWindow (hWnd)
    
    Set WindowElement = AutomationObj.ElementFromHandle(ByVal hWnd)
    Dim iCnd As IUIAutomationCondition
    Do
        DoEvents
        Sleep 1&
        Set iCnd = AutomationObj.CreatePropertyCondition(UIA_NamePropertyId, "�ۑ�")
    Loop While iCnd Is Nothing
    
    Set saveButton = WindowElement.FindFirst(TreeScope_Subtree, iCnd)
    Dim InvokePattern As IUIAutomationInvokePattern
    Set InvokePattern = saveButton.GetCurrentPattern(UIA_InvokePatternId)
    
    Sleep 1000
    InvokePattern.Invoke
    
    pushSaveButton = True

End Function


