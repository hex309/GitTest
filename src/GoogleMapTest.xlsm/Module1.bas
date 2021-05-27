Attribute VB_Name = "Module1"
Option Explicit
'http://www.excel.studio-kazu.jp/kw/20150429195616.html

'//--------------------------------------
Sub getDistanceAndDulationAPI()
    '//--------------------------------------
    Dim i As Long
    Dim xDoc As New MSXML2.DOMDocument '// [参照設定] で "Microsoft XML, version x.0" にチェック。
    Dim txtURL As String
    For i = 5 To Cells(Rows.Count, "C").End(xlUp).Row
        txtURL = "http://maps.googleapis.com/maps/api/distancematrix/xml?" _
            & "language=ja" _
            & "&origins=" & EncodeURL(Cells(i, "B").Text) _
            & "&destinations=" & EncodeURL(Cells(i, "C").Text) _
            & "&avoid=highways"
'        Cells(i, "F").Value = txtURL
        xDoc.async = False
        xDoc.Load txtURL
        If xDoc.SelectNodes("/DistanceMatrixResponse/status")(0).Text = "OK" Then
            Cells(i, "D").Value = xDoc.SelectNodes("/DistanceMatrixResponse/row/element/duration/text")(0).Text
            Cells(i, "E").Value = xDoc.SelectNodes("/DistanceMatrixResponse/row/element/distance/text")(0).Text
        End If
    Next
End Sub
#If VBA7 Then
'//--------------------------------------
Private Function EncodeURL(ByVal txt As String) As String
    '//--------------------------------------
    Dim objDoc As Object
    Dim objEelm As Object
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, "'", "\'")
    Set objDoc = CreateObject("HtmlFile")
    Set objEelm = objDoc.CreateElement("Span")
    objEelm.SetAttribute "id", "result"
    objDoc.AppendChild objEelm
    objDoc.ParentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & txt & "');", "JScript"
    EncodeURL = objEelm.innerText
End Function
 #Else
 '//--------------------------------------
 Private Function EncodeURL(ByVal txt As String) As String
 '//--------------------------------------
    With CreateObject("ScriptControl")
        .Language = "JScript"
        EncodeURL = .CodeObject.encodeURIComponent(txt)
    End With
 End Function
 #End If

