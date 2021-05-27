Attribute VB_Name = "Module3"
Option Explicit

Private Sub Test()
    Dim strm As New ADODB.Stream
    Dim acbPath As String
    Dim str As String
    Dim strData As String
    Dim objIE As Object
  

    Set objIE = CreateObject("InternetExplorer.Application")
    '↓IEを表示する設定
    objIE.Visible = True
    '↓開いているブックのパスを取得
    acbPath = ActiveWorkbook.Path

    '-----地図情報作成(mapdata.jsに書き込む内容)

    '1 地図を表示するズームを数値で設定
    '2地図の中心にする住所を設定
    '3アイコンを表示する住所を設定
 
    strData = "var mapzoom = " & 15 & ";" & vbCrLf & _
        "var places = [" & vbCrLf & _
        "['" & ActiveCell.Text & "','" & ActiveCell.Offset(, 1).Text & "'" & ", 0]," & vbCrLf & _
        "['" & ActiveCell.Text & "','" & ActiveCell.Offset(, 1).Text & "', 1]," & vbCrLf & _
        "];" & vbCrLf

    '------地図情報作成終了

 
    With strm   'strData (地図情報)をmapdata.jsに保存
        .Charset = "UTF-8"  '文字コードの指定
        .Open
        .WriteText strData  '保存する内容
        .SaveToFile acbPath & "\mapdata.js", adSaveCreateOverWrite  '保存先の設定、保存の設定
        .Close: Set strm = Nothing
    End With
  
    '↓IEに表示するHTMLのパスを設定
    str = acbPath & "\index.html"
  
    '↓IEのアドレスを指定
    objIE.navigate str

    'IEが完全表示されるまで待機
    Do While objIE.Busy = True Or objIE.readyState <> 4
        DoEvents
    Loop
    
    Set objIE = Nothing
        
End Sub


