Attribute VB_Name = "WebBrowserSample"
Option Explicit

'https://www.ka-net.org/blog/?p=13587
'���I��WebBrowser�R���g���[����ǉ�����Web�y�[�W�\�����s���T���v��
'��[VBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������]�v�L��
Public Sub Sample()
  Dim frmBrowser As Object
  Const ComponentName = "UserForm1"
  Const ControlName = "WebView"
  Const CtrlWidth = 640
  Const CtrlHeight = 480
   
  '���O��UserForm�폜
  On Error Resume Next
  With Application.VBE.ActiveVBProject.VBComponents
    .Remove .Item(ComponentName)
  End With
  On Error GoTo 0
   
  'UserForm(VBIDE.VBComponent)�ǉ�
  With Application.VBE.ActiveVBProject.VBComponents.Add(3) 'vbext_ct_MSForm
    .Name = ComponentName
    .Properties("Caption").Value = "WebBrowser"
    .Properties("BackColor").Value = &HFFFFFF
    .Properties("Width").Value = CtrlWidth
    .Properties("Height").Value = CtrlHeight
    .Properties("StartUpPosition").Value = 1
    .Properties("ShowModal").Value = False
     
    'WebBrowser(MSForms.Control)�ǉ�
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
   
  'WebBrowser����
  With frmBrowser.Controls(ControlName)
    '�C�x�ߒ��x�Ƀw�b�_�[��User-Agent�ǉ�
    .Navigate2 _
      Url:="https://www.google.com/?hl=ja", _
      Headers:="User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 Edg/90.0.818.66"
    Do While .Busy = True Or .ReadyState <> 4
      DoEvents
    Loop
    .Document.getElementsByName("q")(0).Value = "���S�Ҕ��Y�^"
    .Document.getElementsByName("btnK")(0).Click
  End With
End Sub
