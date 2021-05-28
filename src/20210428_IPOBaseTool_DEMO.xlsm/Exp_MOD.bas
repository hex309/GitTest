Attribute VB_Name = "Exp_MOD"
Option Explicit
'���o�͌n���W���[��

Public Sub Exp_CSV(Optional ByVal str_Stn As String = "Exp_CSV")
'��CSV�o�� �ėp �f�t�H���g=Exp_CSV(�ꎞ�V�[�g)
    '�\�����f�[�^��TMP�V�[�g��CSV�t�@�C��
    Dim L_Ws As Worksheet
    Dim R_Ws As Worksheet
    Dim i, r, eRow As Long
    Dim str_eCol As String
    Dim OutPath As Variant
    Dim Exp_Fn As String
    
    Application.ScreenUpdating = False
    Set R_Ws = Sheets("�Ǘ��\�ҏW�o�^")
    If R_Ws.Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = "" Then
        MsgBox "�o�͂���f�[�^������܂���", 16
        End
    End If
    Set L_Ws = Sheets(str_Stn)
    eRow = R_Ws.Cells(Rows.Count, 4).End(xlUp).Row
    str_eCol = R_Ws.Range("B7").End(xlToRight).Address
    str_eCol = Replace(str_eCol, "7", eRow)
    str_eCol = Replace(str_eCol, "$", "")
    Call Opn_ExlRs("�Ǘ��\�ҏW�o�^$B7:" & str_eCol, "T_2")
    With L_Ws
        .Unprotect
        .Cells.ClearContents
        R_Ws.Range("B7:" & str_eCol).Copy
        .Range("A2").PasteSpecial Paste:=xlValues
        .Range("2:80000").Delete
        .Range("A2").CopyFromRecordset Exl_Rs
        .Visible = True
        .Copy
    End With
    Call Dis_Exl_Rs
    Exp_Fn = Format(Now, "YYMMDDHHMMSS")
    Exp_Fn = Exp_Fn & "_Export_Sample"
    OutPath = Application.GetSaveAsFilename(InitialFileName:=Exp_Fn _
    , FileFilter:="CSV�t�@�C��(*.csv),*.csv", FilterIndex:=1, Title:="�ۑ���̎w��")
    If OutPath <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=OutPath, FileFormat:=xlCSV
        ActiveWorkbook.Close
        MsgBox "CSV�t�@�C���o�͂��������܂����I", vbInformation
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If
    L_Ws.Visible = False
    Call St_Lock
   
End Sub

Public Sub Run_Exp_Backup()
'���o�b�N�A�b�vCSV�쐬 T_KANRI(Access)��TMP��CSV�t�@�C��
    Dim L_Ws As Worksheet
    Dim OutPath As Variant
    Dim Exp_Fn As String
    
    Set L_Ws = Sheets("Exp_BackUp")
    Call Opn_AcRs("T_KANRI", "T_1")
    With L_Ws
        .Unprotect
        .Range("3:80000").Delete
        .Range("A3").CopyFromRecordset Ac_Rs
    End With
    L_Ws.Visible = True
    L_Ws.Copy
    Exp_Fn = Format(Now, "YYMMDDHHMMSS")
    Exp_Fn = Exp_Fn & "_BackUp"
    Application.ScreenUpdating = False
    OutPath = Application.GetSaveAsFilename(InitialFileName:=Exp_Fn _
    , FileFilter:="CSV�t�@�C��(*.csv),*.csv", FilterIndex:=1, Title:="�ۑ���̎w��")
    If OutPath <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=OutPath, FileFormat:=xlCSV
        ActiveWorkbook.Close
        MsgBox "�o�b�N�A�b�v�t�@�C���o�͂��������܂����I", vbInformation
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If
    L_Ws.Visible = False
 
End Sub

