Attribute VB_Name = "��O_���Ϗo��"
Option Explicit

Sub ��O_���Ϗo�̓��W���[��()

'    Call ���ϊm�F�����o�^�A�b�v���[�h�pCSV�t�@�C���o��
    Call ��O_���ϊm�F�����o�^_��������

End Sub


'Sub ���ϊm�F�����o�^�A�b�v���[�h�pCSV�t�@�C���o��()
'
'    Dim �V�[�g��    As String
'
'    Call ����_���ϓo�^�A�b�v���[�h�p.���ϊm�F�����o�^�A�b�v���[�h�pCSV�t�@�C���쐬(Pub���σV�[�g��)
'
'End Sub

Sub ��O_���ϊm�F�����o�^_��������()
   
    ThisWorkbook.Sheets(WSNAME_WARIKOMI).Range("H20:H100").ClearContents
    
    Dim buf As String
    Dim tmp As Variant
    
    Open ThisWorkbook.Path & "\���A�b�v���[�h�p�t�@�C��\���ϊm�F�����o�^(�A�b�v���[�h�p).csv" For Input As #1
    
    Dim i As Long: i = 1
    Dim z As Long: z = 0
    Do Until EOF(1)
        Line Input #1, buf
        tmp = Split(buf, ",")

        If i >= 6 Then
            If tmp(25) = 1 Then ThisWorkbook.Sheets(WSNAME_WARIKOMI).Cells(14 + i, 8).Value = "��"  '�����@�Ώ�
            If tmp(26) = 1 Then ThisWorkbook.Sheets(WSNAME_WARIKOMI).Cells(14 + i + 1, 8).Value = "��"  '���Ɩ@�Ώ�
            i = i + 1
        End If
        If 14 + i = 39 Then i = i + 2
        i = i + 1
    Loop
    
    Close #1

End Sub

