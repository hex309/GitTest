Attribute VB_Name = "����_�ʒm�o�["
Option Explicit

Public Sub ieDownloadFileNbOrDlg(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)
    
    Call ieWaitCheck

    Dim hIE As Long
    Dim SavePath As String

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        hIE = oIE.hWnd
        SavePath = Pub��ƃt�H���_�p�X
        bFLG = True
   
continue1:
        
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
    
    Dim uiAuto As CUIAutomation
    Dim elmIE As IUIAutomationElement
    Dim elmNotificationBar As IUIAutomationElement
    Dim elmSaveSplitButton As IUIAutomationElement
    Dim elmSaveDropDownButton As IUIAutomationElement
    Dim elmSaveMenu As IUIAutomationElement
    Dim elmSaveMenuItem As IUIAutomationElement
    Dim elmIEDialog As IUIAutomationElement
    Dim elmSaveAsButton As IUIAutomationElement
    Dim elmSaveAsWindow As IUIAutomationElement
    Dim elmFileNameEdit As IUIAutomationElement
    Dim elmSaveButton As IUIAutomationElement
    Dim elmNotificationText As IUIAutomationElement
    Dim elmCloseButton As IUIAutomationElement
    Dim iptn As IUIAutomationInvokePattern
    Dim vptn As IUIAutomationValuePattern
    Const ROLE_SYSTEM_BUTTONDROPDOWN = &H38&
   
    Set uiAuto = New CUIAutomation
    Set elmIE = uiAuto.ElementFromHandle(ByVal hIE)
     
  '�t�@�C�������O�ɍ폜
'  With CreateObject("Scripting.FileSystemObject")
'    If .FileExists(SaveFilePath) Then .DeleteFile SaveFilePath, True
'  End With
   
  Do
    '[�ʒm�o�[]�擾
    Set elmNotificationBar = _
      GetElement(uiAuto, _
                 elmIE, _
                 UIA_AutomationIdPropertyId, _
                 "IENotificationBar", _
                 UIA_ToolBarControlTypeId)
     
    '[Internet Explorer]�_�C�A���O((�t�@�C����) �ōs�������I��ł�������)�擾
    Set elmIEDialog = _
      GetElement(uiAuto, _
                 elmIE, _
                 UIA_NamePropertyId, _
                 "Internet Explorer", _
                 UIA_WindowControlTypeId)
    DoEvents
  Loop Until (Not elmNotificationBar Is Nothing) Or _
             (Not elmIEDialog Is Nothing)
   
  '***** �ʒm�o�[���삱������ *****
  If Not elmNotificationBar Is Nothing Then
    '[�ۑ�]�X�v���b�g�{�^���擾
    Set elmSaveSplitButton = _
      GetElement(uiAuto, _
                 elmNotificationBar, _
                 UIA_NamePropertyId, _
                 "�ۑ�", _
                 UIA_SplitButtonControlTypeId)
    If elmSaveSplitButton Is Nothing Then GoTo fin
     
    '[�ۑ�]�h���b�v�_�E���擾
    Set elmSaveDropDownButton = _
      GetElement(uiAuto, _
                 elmNotificationBar, _
                 UIA_LegacyIAccessibleRolePropertyId, _
                 ROLE_SYSTEM_BUTTONDROPDOWN, _
                 UIA_SplitButtonControlTypeId)
    If elmSaveDropDownButton Is Nothing Then GoTo fin
     
    '[�ۑ�]�h���b�v�_�E������ -> [���O��t���ĕۑ�(A)]�{�^������
    Set iptn = elmSaveDropDownButton.GetCurrentPattern(UIA_InvokePatternId)
    Do
      iptn.Invoke
      Set elmSaveMenu = _
        GetElement(uiAuto, _
                   uiAuto.GetRootElement, _
                   UIA_ClassNamePropertyId, _
                   "#32768", _
                   UIA_MenuControlTypeId)
      DoEvents
    Loop While elmSaveMenu Is Nothing
    Set elmSaveMenuItem = _
      GetElement(uiAuto, _
                 elmSaveMenu, _
                 UIA_NamePropertyId, _
                 "���O��t���ĕۑ�(A)", _
                 UIA_MenuItemControlTypeId)
    If elmSaveMenuItem Is Nothing Then GoTo fin
    Set iptn = elmSaveMenuItem.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke
  End If
  '***** �ʒm�o�[���삱���܂� *****
   
  '***** Internet Explorer�_�C�A���O���삱������ *****
  If Not elmIEDialog Is Nothing Then
    Set elmSaveAsButton = _
      GetElement(uiAuto, _
                 elmIEDialog, _
                 UIA_NamePropertyId, _
                 "���O��t���ĕۑ�(A)", _
                 UIA_ButtonControlTypeId)
    If elmSaveAsButton Is Nothing Then GoTo fin
    Set iptn = elmSaveAsButton.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke
  End If
  '***** Internet Explorer�_�C�A���O���삱���܂� *****
   
  If (elmNotificationBar Is Nothing) And (elmIEDialog Is Nothing) Then GoTo fin
   
  '***** ���O��t���ĕۑ����삱������ *****
  '[���O��t���ĕۑ�]�_�C�A���O�擾
  Do
    Set elmSaveAsWindow = _
      GetElement(uiAuto, _
                 uiAuto.GetRootElement, _
                 UIA_NamePropertyId, _
                 "���O��t���ĕۑ�", _
                 UIA_WindowControlTypeId)
    DoEvents
  Loop While elmSaveAsWindow Is Nothing
   
  '[�t�@�C����]���擾 -> �t�@�C���p�X����
  Set elmFileNameEdit = _
    GetElement(uiAuto, _
               elmSaveAsWindow, _
               UIA_NamePropertyId, _
               "�t�@�C����:", _
               UIA_EditControlTypeId)
  If elmFileNameEdit Is Nothing Then GoTo fin
  Set vptn = elmFileNameEdit.GetCurrentPattern(UIA_ValuePatternId)

  vptn.SetValue SavePath & vptn.CurrentValue
  
'  vptn.SetValue SaveFilePath
   
  '[�ۑ�(S)]�{�^������
  Set elmSaveButton = _
    GetElement(uiAuto, _
               elmSaveAsWindow, _
               UIA_NamePropertyId, _
               "�ۑ�(S)", _
               UIA_ButtonControlTypeId)
  If elmSaveButton Is Nothing Then GoTo fin
  Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
  iptn.Invoke
  '***** ���O��t���ĕۑ����삱���܂� *****
   
  '***** �_�E�����[�h�����҂��������� *****
  If elmNotificationBar Is Nothing Then
    '[�ʒm�o�[]�擾
    Do
      Set elmNotificationBar = _
        GetElement(uiAuto, _
                   elmIE, _
                   UIA_AutomationIdPropertyId, _
                   "IENotificationBar", _
                   UIA_ToolBarControlTypeId)
      DoEvents
    Loop While elmNotificationBar Is Nothing
  End If
   
  '[�ʒm�o�[�̃e�L�X�g]�擾
  Set elmNotificationText = _
    GetElement(uiAuto, _
               elmNotificationBar, _
               UIA_NamePropertyId, _
               "�ʒm�o�[�̃e�L�X�g", _
               UIA_TextControlTypeId)
  If elmNotificationText Is Nothing Then GoTo fin
   
  '[����]�{�^���擾
  Set elmCloseButton = _
    GetElement(uiAuto, _
               elmNotificationBar, _
               UIA_NamePropertyId, _
               "����", _
               UIA_ButtonControlTypeId)
  If elmCloseButton Is Nothing Then GoTo fin
   
  Do
    DoEvents
  Loop Until InStr( _
    elmNotificationText.GetCurrentPropertyValue(UIA_ValueValuePropertyId), _
    "�_�E�����[�h���������܂���") > 0
   
  '[����]�{�^������
  Set iptn = elmCloseButton.GetCurrentPattern(UIA_InvokePatternId)
  iptn.Invoke
  '***** �_�E�����[�h�����҂������܂� *****
   
  Exit Sub
fin:
  MsgBox "���������s���܂����B", vbCritical + vbSystemModal
End Sub

Public Function GetElement(ByVal uiAuto As CUIAutomation, ByVal elmParent As IUIAutomationElement, ByVal propertyId As Long, ByVal propertyValue As Variant, Optional ByVal ctrlType As Long = 0) As IUIAutomationElement
                            
    Dim cndFirst As IUIAutomationCondition
    Dim cndSecond As IUIAutomationCondition
    
    Set cndFirst = uiAuto.CreatePropertyCondition(propertyId, propertyValue)
    
    If ctrlType <> 0 Then
        Set cndSecond = uiAuto.CreatePropertyCondition(UIA_ControlTypePropertyId, ctrlType)
        Set cndFirst = uiAuto.CreateAndCondition(cndFirst, cndSecond)
    End If
    
    Set GetElement = elmParent.FindFirst(TreeScope_Subtree, cndFirst)
    
End Function
  

