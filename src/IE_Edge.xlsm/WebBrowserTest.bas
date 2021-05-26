Attribute VB_Name = "WebBrowserTest"
Option Explicit

Public Sub Sample()
    frmBrowser.Show
   
    'WebBrowser操作
    With frmBrowser.Controls("WebBrowser")
        '気休め程度にヘッダーにUser-Agent追加
        'https://www.yahoo.co.jp/
        'https://www.google.com/?hl=ja
        .Navigate2 _
            URL:="https://www.google.com/?hl=ja", _
            Headers:="User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36 Edg/90.0.818.66"
        Do While .Busy = True Or .ReadyState <> 4
            DoEvents
        Loop
        .Document.getElementsByName("q")(0).Value = "初心者備忘録"
        .Document.getElementsByName("btnK")(0).Click
    End With
End Sub



