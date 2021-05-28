Attribute VB_Name = "Exp_MOD"
Option Explicit
'★出力系モジュール

Public Sub Exp_CSV(Optional ByVal str_Stn As String = "Exp_CSV")
'★CSV出力 汎用 デフォルト=Exp_CSV(一時シート)
    '表示中データ⇒TMPシート⇒CSVファイル
    Dim L_Ws As Worksheet
    Dim R_Ws As Worksheet
    Dim i, r, eRow As Long
    Dim str_eCol As String
    Dim OutPath As Variant
    Dim Exp_Fn As String
    
    Application.ScreenUpdating = False
    Set R_Ws = Sheets("管理表編集登録")
    If R_Ws.Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = "" Then
        MsgBox "出力するデータがありません", 16
        End
    End If
    Set L_Ws = Sheets(str_Stn)
    eRow = R_Ws.Cells(Rows.Count, 4).End(xlUp).Row
    str_eCol = R_Ws.Range("B7").End(xlToRight).Address
    str_eCol = Replace(str_eCol, "7", eRow)
    str_eCol = Replace(str_eCol, "$", "")
    Call Opn_ExlRs("管理表編集登録$B7:" & str_eCol, "T_2")
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
    , FileFilter:="CSVファイル(*.csv),*.csv", FilterIndex:=1, Title:="保存先の指定")
    If OutPath <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=OutPath, FileFormat:=xlCSV
        ActiveWorkbook.Close
        MsgBox "CSVファイル出力が完了しました！", vbInformation
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If
    L_Ws.Visible = False
    Call St_Lock
   
End Sub

Public Sub Run_Exp_Backup()
'★バックアップCSV作成 T_KANRI(Access)⇒TMP⇒CSVファイル
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
    , FileFilter:="CSVファイル(*.csv),*.csv", FilterIndex:=1, Title:="保存先の指定")
    If OutPath <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs FileName:=OutPath, FileFormat:=xlCSV
        ActiveWorkbook.Close
        MsgBox "バックアップファイル出力が完了しました！", vbInformation
    Else
        ActiveWorkbook.Close SaveChanges:=False
    End If
    L_Ws.Visible = False
 
End Sub

