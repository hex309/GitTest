Attribute VB_Name = "����_���ϓo�^�A�b�v���[�h�p"
Option Explicit

Sub ���ϊm�F�����o�^�A�b�v���[�h�pCSV�t�@�C���쐬(ByVal ���ϓo�^�V�[�g�� As String)
    
    Dim SrcXLFName As String
    Dim SrcSHTName As String
    
    Pub�u�b�N�� = "�y�����z���ϊm�F���ꗗ.xlsx"
    
    SrcXLFName = ThisWorkbook.Path & "\" & Pub�u�b�N��
    SrcSHTName = ���ϓo�^�V�[�g�� & "$"
    
    Dim Cn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    
    Set Cn = New ADODB.Connection
    Set Rs = New ADODB.Recordset
    
    Cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                        "Data Source=" & SrcXLFName & ";" & _
                        "Extended Properties=""Excel 12.0;HDR=No;IMEX=1;Readonly=False"""
    
    Set Rs = Cn.Execute("SELECT * FROM [" & SrcSHTName & "]")

    Dim CSVFile As String
    
    CSVFile = ThisWorkbook.Path & "\���A�b�v���[�h�p�t�@�C��\���ϊm�F�����o�^(�A�b�v���[�h�p).csv"
        
    If Rs.RecordCount = 0 Then
        MsgBox "ADODB:���R�[�h��������܂���", vbCritical
        End
    Else
        
        Open CSVFile For Output As #1
        
        Dim i As Long: i = 1
        Dim j As Long: j = 0
        Do Until Rs.EOF
            
            If i = 3 Then '�s��
'                Print #1, Rs.Fields(0).Value & ",";
'                Print #1, Rs.Fields(1).Value & ",";
'                Print #1, Rs.Fields(2).Value & ",";
'                Print #1, Rs.Fields(3).Value & ",";
'                Print #1, Rs.Fields(4).Value & ",";
'                Print #1, CDate(Rs.Fields(5).Value) & ",";
'                Print #1, Rs.Fields(6).Value & ",";
'                Print #1, Rs.Fields(7).Value & ",";
'                Print #1, Rs.Fields(8).Value & ",";
'                Print #1, Rs.Fields(8).Value & ",";
'                Print #1, Rs.Fields(10).Value & ",";
'                Print #1, Rs.Fields(11).Value & ",";
'                Print #1, Rs.Fields(12).Value & ",";
'                Print #1, Rs.Fields(13).Value & ",";
'                Print #1, Rs.Fields(14).Value & ",";
'                Print #1, Rs.Fields(15).Value & ",";
'                Print #1, Rs.Fields(16).Value & ",";
'                Print #1, Rs.Fields(17).Value & ",";
'                Print #1, Rs.Fields(18).Value & ",";
'                Print #1, Rs.Fields(19).Value & ",";
'                Print #1, Rs.Fields(20).Value & ",";
'                Print #1, Rs.Fields(21).Value & ",";
'                Print #1, Rs.Fields(22).Value & ",";
'                Print #1, Rs.Fields(23).Value & ","
                
                For j = 0 To Rs.Fields.Count - 1
                    If j <> Rs.Fields.Count - 1 Then
                        
                        If j = 4 Then
                            On Error Resume Next 'orz
                            If PubClsLucasAuth.LucasID = "" Then
                                Print #1, Replace(IIf(IsNull(Rs.Fields(j).Value), "", Rs.Fields(j).Value), ",", "") & ",";
                            Else
                                Print #1, PubClsLucasAuth.LucasID & ",";
                            End If
                            On Error GoTo 0
                        ElseIf j = 5 Then
                            Print #1, CDate(Rs.Fields(5).Value) & ",";  '���ϗL������
                        ElseIf j = 11 Then
                            Print #1, "" & ","; '���ϑO�����(CSV�t�@�C���ɂ̓u�����N���w�肷��B�W���u�V�[�g���œ��͂����邽��)
                        Else
                            Print #1, Replace(IIf(IsNull(Rs.Fields(j).Value), "", Rs.Fields(j).Value), ",", "") & ",";
                        End If
                    
                    Else
                        Print #1, Rs.Fields(j).Value & ","
                    End If
                Next
            Else
                For j = 0 To Rs.Fields.Count - 1
                    If j <> Rs.Fields.Count - 1 Then
                        Print #1, Replace(IIf(IsNull(Rs.Fields(j).Value), "", Rs.Fields(j).Value), ",", "") & ",";
                    Else
                        Print #1, Replace(IIf(IsNull(Rs.Fields(j).Value), "", Rs.Fields(j).Value), ",", "") & ","
                    End If
                Next
            End If
            i = i + 1
            Rs.MoveNext
        Loop
    
        Close #1
    
    End If
        
    Rs.Close: Set Rs = Nothing
    Cn.Close: Set Cn = Nothing

End Sub

'Sub Sample2()
'
'  Dim WB_PATH As String
'  Dim WB_NAME As String
'  Dim WS_NAME As String
'
'  Dim ROW As Long
'  Dim COL As Long
'
'  WB_PATH = ThisWorkbook.Path & "\"  '�u�b�N�p�X(�p�X�̏I����\)
'  WB_NAME = "�y�����z���ϊm�F���ꗗ.xlsx" '�u�b�N��
'  WS_NAME = "POS1SiR�ێ�"      '�V�[�g��
'
'  ROW = 3 '�s�ԍ�
'  COL = 11  '��ԍ�
'
'  Dim STR As String
'
'  STR = "'" & WB_PATH & "[" & WB_NAME & "]" & WS_NAME & "'!R" & ROW & "C" & COL & ""
'
'  MsgBox ExecuteExcel4Macro(STR)
'
'End Sub




