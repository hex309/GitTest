Attribute VB_Name = "Module3"
Option Explicit

Private Sub Test()
    Dim strm As New ADODB.Stream
    Dim acbPath As String
    Dim str As String
    Dim strData As String
    Dim objIE As Object
  

    Set objIE = CreateObject("InternetExplorer.Application")
    '��IE��\������ݒ�
    objIE.Visible = True
    '���J���Ă���u�b�N�̃p�X���擾
    acbPath = ActiveWorkbook.Path

    '-----�n�}���쐬(mapdata.js�ɏ������ޓ��e)

    '1 �n�}��\������Y�[���𐔒l�Őݒ�
    '2�n�}�̒��S�ɂ���Z����ݒ�
    '3�A�C�R����\������Z����ݒ�
 
    strData = "var mapzoom = " & 15 & ";" & vbCrLf & _
        "var places = [" & vbCrLf & _
        "['" & ActiveCell.Text & "','" & ActiveCell.Offset(, 1).Text & "'" & ", 0]," & vbCrLf & _
        "['" & ActiveCell.Text & "','" & ActiveCell.Offset(, 1).Text & "', 1]," & vbCrLf & _
        "];" & vbCrLf

    '------�n�}���쐬�I��

 
    With strm   'strData (�n�}���)��mapdata.js�ɕۑ�
        .Charset = "UTF-8"  '�����R�[�h�̎w��
        .Open
        .WriteText strData  '�ۑ�������e
        .SaveToFile acbPath & "\mapdata.js", adSaveCreateOverWrite  '�ۑ���̐ݒ�A�ۑ��̐ݒ�
        .Close: Set strm = Nothing
    End With
  
    '��IE�ɕ\������HTML�̃p�X��ݒ�
    str = acbPath & "\index.html"
  
    '��IE�̃A�h���X���w��
    objIE.navigate str

    'IE�����S�\�������܂őҋ@
    Do While objIE.Busy = True Or objIE.readyState <> 4
        DoEvents
    Loop
    
    Set objIE = Nothing
        
End Sub


