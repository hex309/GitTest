Attribute VB_Name = "UI_MOD"
Option Explicit
'★UI表示周り系モジュール

Public Sub OP_form0()
'★管理表ID入力フォーム起動
    Dim eRow As Long
    Dim Ans As String
    
    eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
    ActiveSheet.Unprotect
    If Cells(eRow, 7).Value = "" Then
        Cells(eRow, 7).Select
        Ans = MsgBox("外部カラムIDが設定されていない管理表IDがあります" & vbCrLf & _
        "まだ設定されていない外部カラムIDを設定してください" & vbCrLf & _
        "未設定の管理表IDを破棄して続けますか？", vbYesNo + vbInformation, "外部カラムIDが未設定のIDがあります")
        If Ans = vbNo Then
            Call St_Lock
            End
        ElseIf Ans = vbYes Then
            Cells(eRow, 5).ClearContents
        End If
    End If
    Call St_Lock
    UF_0.Show

End Sub

Public Sub OP_form1()
'★外部ID入力フォーム起動
    Dim Ws As Worksheet
    Dim eRow As Long
    
    Set Ws = Sheets("カラム設定")
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "管理表カラムIDを設定してから行ってください", 16, "管理表カラムID未入力エラー"
            Exit Sub
        End If
    End With
    UF_1.Show

End Sub

Public Sub ALL_Reset()
'★UIリセット
    Call St_AllUnvis
    With Sheets("インポート")
        .Unprotect
        .Shapes("Fil_1").Visible = False
        .Shapes("Gr_1").Visible = False
    End With

End Sub

Public Sub St_AllUnvis()
'★ホームシート以外のシート非表示
    Dim Stc As Long
    Dim i As Long
    
    Stc = ThisWorkbook.Sheets.Count
    Sheets("ホーム").Visible = True
    For i = 2 To Stc
        Sheets(i).Visible = False
    Next i

End Sub

Public Sub vis_UISt()
'★ホーム画面シート表示
    Call vis_St("ホーム")
    
End Sub

Public Sub vis_ImportSt()
'★インポート画面シート表示
    Call vis_St("インポート")
    ActiveSheet.Unprotect
    ActiveSheet.Range("C7").ClearContents
    Call St_Lock

End Sub

Public Sub vis_CSVImportSt()
'★インポート画面シート表示
    Call vis_St("CSVインポート")
    ActiveSheet.Unprotect
    ActiveSheet.Range("C7").ClearContents
    Call St_Lock

End Sub

Public Sub vis_KANRISt()
'★管理表編集登録シート表示
    Application.ScreenUpdating = False
    Call vis_St("管理表編集登録")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call Re_KANRI
    Call Re_Scrl
    ActiveSheet.Range("E10").Select

End Sub

Public Sub vis_RegNewIDSt()
'★管理表新規登録シート表示
    Call vis_St("管理表新規登録")
    Sheets("管理表新規登録").Unprotect
    Sheets("管理表新規登録").Range("D6").ClearContents
    Call St_Lock
    
End Sub

Public Sub vis_CosKANRISt()
'★カスタム編集登録シート表示
    Application.ScreenUpdating = False
    Call vis_St("管理表編集登録")
    Call Re_CosKANRI
    Call Re_Scrl
    
End Sub

Public Sub vis_KANRIvewSt()
'★管理表出力ビューシート表示
    Application.ScreenUpdating = False
    Call vis_St("管理表出力ビュー")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call Run_Douki_KANRIvew
    Call Re_Scrl
    Call St_Lock

End Sub

Public Sub vis_CostumvewSt()
'★カスタムビューシート表示
    Application.ScreenUpdating = False
    Call vis_St("カスタムビュー")
    ActiveSheet.Unprotect
    ActiveSheet.Rows(4).ClearContents
    Call St_Lock
'    Call Run_Search_Costumvew
  
End Sub

Public Sub vis_T_GAIBSt()
'★外部データシート表示
    Call vis_St("外部データ")
    Call Run_Douki_GAIB
    Call Re_Scrl
    Call Clear_SearchKey_G
    
End Sub

Public Sub vis_SETTEISt()
'★設定シート表示
    Call vis_St("設定")

End Sub

Public Sub vis_DBSETTEISt()
'★データベース設定シート表示
    Call vis_St("データベース設定")

End Sub
Public Sub vis_SETDirectSt()
'★ディレクトリ設定シート表示
    Call vis_St("ディレクトリ設定")

End Sub

Public Sub vis_SETFildsSt()
'★管理表フィールド設定シート表示
    Call vis_St("管理表フィールド設定")

End Sub

Public Sub vis_RAKRangSt()
'★シート範囲設定シート表示
    Call vis_St("外部データシート範囲設定")

End Sub

Public Sub vis_TOESET1St()
'★カラム設定1シート表示
    Call vis_St("カラム設定")
    Call Re_Scrl
    
End Sub

Public Function vis_St(ByVal Op_St As String)
'☆指定シートの表示⇒開いているシートの非表示
    With Sheets(Op_St)
        .Visible = True
        ActiveSheet.Visible = False
    End With
    Call St_Lock
    
End Function

Sub St_Lock()
'☆シートロック
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
End Sub


Public Sub Re_Scrl()
'☆スクロールリセット
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
End Sub
