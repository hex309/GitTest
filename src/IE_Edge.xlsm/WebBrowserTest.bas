Attribute VB_Name = "WebBrowserTest"
Option Explicit

Public Sub Sample()
    frmBrowser.Show
   
    'WebBrowser����
    With frmBrowser.Controls("WebBrowser")
        '�C�x�ߒ��x�Ƀw�b�_�[��User-Agent�ǉ�
        'https://www.yahoo.co.jp/
        'https://www.google.com/?hl=ja
        .Navigate2 _
            URL:="https://www.google.com/?hl=ja", _
            Headers:="User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 Edg/90.0.818.66"
        Do While .Busy = True Or .ReadyState <> 4
            DoEvents
        Loop
        .Document.getElementsByName("q")(0).Value = "���S�Ҕ��Y�^"
        .Document.getElementsByName("btnK")(0).Click
    End With
End Sub



