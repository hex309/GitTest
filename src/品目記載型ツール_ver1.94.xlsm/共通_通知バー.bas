Attribute VB_Name = "共通_通知バー"
Option Explicit

Public Sub ieDownloadFileNbOrDlg(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)
    
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

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        hIE = oIE.hWnd
        SavePath = Pub作業フォルダパス
        bFLG = True
   
continue1:
        
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
    
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
     
  'ファイルを事前に削除
'  With CreateObject("Scripting.FileSystemObject")
'    If .FileExists(SaveFilePath) Then .DeleteFile SaveFilePath, True
'  End With
   
  Do
    '[通知バー]取得
    Set elmNotificationBar = _
      GetElement(uiAuto, _
                 elmIE, _
                 UIA_AutomationIdPropertyId, _
                 "IENotificationBar", _
                 UIA_ToolBarControlTypeId)
     
    '[Internet Explorer]ダイアログ((ファイル名) で行う操作を選んでください)取得
    Set elmIEDialog = _
      GetElement(uiAuto, _
                 elmIE, _
                 UIA_NamePropertyId, _
                 "Internet Explorer", _
                 UIA_WindowControlTypeId)
    DoEvents
  Loop Until (Not elmNotificationBar Is Nothing) Or _
             (Not elmIEDialog Is Nothing)
   
  '***** 通知バー操作ここから *****
  If Not elmNotificationBar Is Nothing Then
    '[保存]スプリットボタン取得
    Set elmSaveSplitButton = _
      GetElement(uiAuto, _
                 elmNotificationBar, _
                 UIA_NamePropertyId, _
                 "保存", _
                 UIA_SplitButtonControlTypeId)
    If elmSaveSplitButton Is Nothing Then GoTo fin
     
    '[保存]ドロップダウン取得
    Set elmSaveDropDownButton = _
      GetElement(uiAuto, _
                 elmNotificationBar, _
                 UIA_LegacyIAccessibleRolePropertyId, _
                 ROLE_SYSTEM_BUTTONDROPDOWN, _
                 UIA_SplitButtonControlTypeId)
    If elmSaveDropDownButton Is Nothing Then GoTo fin
     
    '[保存]ドロップダウン押下 -> [名前を付けて保存(A)]ボタン押下
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
                 "名前を付けて保存(A)", _
                 UIA_MenuItemControlTypeId)
    If elmSaveMenuItem Is Nothing Then GoTo fin
    Set iptn = elmSaveMenuItem.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke
  End If
  '***** 通知バー操作ここまで *****
   
  '***** Internet Explorerダイアログ操作ここから *****
  If Not elmIEDialog Is Nothing Then
    Set elmSaveAsButton = _
      GetElement(uiAuto, _
                 elmIEDialog, _
                 UIA_NamePropertyId, _
                 "名前を付けて保存(A)", _
                 UIA_ButtonControlTypeId)
    If elmSaveAsButton Is Nothing Then GoTo fin
    Set iptn = elmSaveAsButton.GetCurrentPattern(UIA_InvokePatternId)
    iptn.Invoke
  End If
  '***** Internet Explorerダイアログ操作ここまで *****
   
  If (elmNotificationBar Is Nothing) And (elmIEDialog Is Nothing) Then GoTo fin
   
  '***** 名前を付けて保存操作ここから *****
  '[名前を付けて保存]ダイアログ取得
  Do
    Set elmSaveAsWindow = _
      GetElement(uiAuto, _
                 uiAuto.GetRootElement, _
                 UIA_NamePropertyId, _
                 "名前を付けて保存", _
                 UIA_WindowControlTypeId)
    DoEvents
  Loop While elmSaveAsWindow Is Nothing
   
  '[ファイル名]欄取得 -> ファイルパス入力
  Set elmFileNameEdit = _
    GetElement(uiAuto, _
               elmSaveAsWindow, _
               UIA_NamePropertyId, _
               "ファイル名:", _
               UIA_EditControlTypeId)
  If elmFileNameEdit Is Nothing Then GoTo fin
  Set vptn = elmFileNameEdit.GetCurrentPattern(UIA_ValuePatternId)

  vptn.SetValue SavePath & vptn.CurrentValue
  
'  vptn.SetValue SaveFilePath
   
  '[保存(S)]ボタン押下
  Set elmSaveButton = _
    GetElement(uiAuto, _
               elmSaveAsWindow, _
               UIA_NamePropertyId, _
               "保存(S)", _
               UIA_ButtonControlTypeId)
  If elmSaveButton Is Nothing Then GoTo fin
  Set iptn = elmSaveButton.GetCurrentPattern(UIA_InvokePatternId)
  iptn.Invoke
  '***** 名前を付けて保存操作ここまで *****
   
  '***** ダウンロード完了待ちここから *****
  If elmNotificationBar Is Nothing Then
    '[通知バー]取得
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
   
  '[通知バーのテキスト]取得
  Set elmNotificationText = _
    GetElement(uiAuto, _
               elmNotificationBar, _
               UIA_NamePropertyId, _
               "通知バーのテキスト", _
               UIA_TextControlTypeId)
  If elmNotificationText Is Nothing Then GoTo fin
   
  '[閉じる]ボタン取得
  Set elmCloseButton = _
    GetElement(uiAuto, _
               elmNotificationBar, _
               UIA_NamePropertyId, _
               "閉じる", _
               UIA_ButtonControlTypeId)
  If elmCloseButton Is Nothing Then GoTo fin
   
  Do
    DoEvents
  Loop Until InStr( _
    elmNotificationText.GetCurrentPropertyValue(UIA_ValueValuePropertyId), _
    "ダウンロードが完了しました") > 0
   
  '[閉じる]ボタン押下
  Set iptn = elmCloseButton.GetCurrentPattern(UIA_InvokePatternId)
  iptn.Invoke
  '***** ダウンロード完了待ちここまで *****
   
  Exit Sub
fin:
  MsgBox "処理が失敗しました。", vbCritical + vbSystemModal
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
  

