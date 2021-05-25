Attribute VB_Name = "WebBrowserSample"
Option Explicit

'https://www.ka-net.org/blog/?p=13587
'動的にWebBrowserコントロールを追加してWebページ表示を行うサンプル
'※[VBA プロジェクト オブジェクト モデルへのアクセスを信頼する]要有効
Public Sub Sample()
  Dim frmBrowser As Object
  Const ComponentName = "UserForm1"
  Const ControlName = "WebView"
  Const CtrlWidth = 640
  Const CtrlHeight = 480
   
  '事前にUserForm削除
  On Error Resume Next
  With Application.VBE.ActiveVBProject.VBComponents
    .Remove .Item(ComponentName)
  End With
  On Error GoTo 0
   
  'UserForm(VBIDE.VBComponent)追加
  With Application.VBE.ActiveVBProject.VBComponents.Add(3) 'vbext_ct_MSForm
    .Name = ComponentName
    .Properties("Caption").Value = "WebBrowser"
    .Properties("BackColor").Value = &HFFFFFF
    .Properties("Width").Value = CtrlWidth
    .Properties("Height").Value = CtrlHeight
    .Properties("StartUpPosition").Value = 1
    .Properties("ShowModal").Value = False
     
    'WebBrowser(MSForms.Control)追加
    With .Designer.Controls.Add("Shell.Explorer.2")
      .Name = ControlName
      .Top = 0
      .Left = 0
      .Width = CtrlWidth
      .Height = CtrlHeight
    End With
  End With
   
  Set frmBrowser = UserForms.Add(ComponentName)
  frmBrowser.Show
   
  'WebBrowser操作
  With frmBrowser.Controls(ControlName)
    '気休め程度にヘッダーにUser-Agent追加
    .Navigate2 _
      Url:="https://www.google.com/?hl=ja", _
      Headers:="User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 Edg/90.0.818.66"
    Do While .Busy = True Or .ReadyState <> 4
      DoEvents
    Loop
    .Document.getElementsByName("q")(0).Value = "初心者備忘録"
    .Document.getElementsByName("btnK")(0).Click
  End With
End Sub

