Attribute VB_Name = "Module2"
Option Explicit
'http://www.excel.studio-kazu.jp/kw/20150429195616.html

 #If VBA7 Then
 Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
 #Else
 Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
 #End If
 '//--------------------------------------
 Sub SearchDistanceAndTime()
 '//--------------------------------------
    Dim i As Long
    Dim j As Long
    Dim objIE As Object
    Set objIE = CreateObject("InternetExplorer.application")
    objIE.Visible = True
    Dim txtURL As String
    For i = 5 To Cells(Rows.Count, "C").End(xlUp).Row
        txtURL = "http://maps.googleapis.com/maps/api/distancematrix/xml?" _
            & "language=ja" _
            & "&origins=" & EncodeURL(Cells(i, "B").Text) _
            & "&destinations=" & EncodeURL(Cells(i, "C").Text) _
            & "&avoid=highways"
        objIE.navigate txtURL
        readyStateWait objIE
        Cells(i, "D").Value = getResultValue(objIE.Document, "distance")
        Cells(i, "E").Value = getResultValue(objIE.Document, "duration")
    Next
    objIE.Quit
 End Sub
 '//--------------------------------------
 Function getResultValue(dom As Object, tagName As String)
 '//--------------------------------------
    Dim res
    res = dom.getElementsByTagName(tagName)(0).getElementsByTagName("text")(0).innerText
    res = Replace(res, "<text>", "")
    getResultValue = Replace(res, "</text>", "")
 End Function
'//--------------------------------------
Sub readyStateWait(objIE As Object)
    '//--------------------------------------
    Do While objIE.readyState <> 4 Or objIE.Busy = True
        DoEvents
    Loop
    Sleep 1000
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

