Attribute VB_Name = "Get_List"
Option Explicit
'★リスト作成系モジュール

Public Sub Get_KANRIColList()
'★管理表データフォーマット　フィールド名のリスト作成
    'カラム設定管理表カラムID入力フォームで使用
   Call Get_FildLis("管理表フィールド設定$B5:GS6", "T_KANRIColList", "管理表ID", 2)

End Sub

Public Sub Get_GAIBColList()
'★外部データフォーマット　フィールド名のリスト作成
    'カラム設定外部カラムID入力フォームで使用
        '255以上を想定　Get_FildLisで取得できるのはMAX255までの為
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim i, r As Long
    Dim CVal As String
    
    With ThisWorkbook
        Set R_Ws = .Sheets("T_GAIBCol")
        Set L_Ws = .Sheets("T_GAIBColList")
        L_Ws.Unprotect
        L_Ws.Range("B2:B500").ClearContents
        For i = 1 To 400
            CVal = R_Ws.Cells(1, i).Value
'            '◆↓改行情報を保持したい場合はコメントアウト
'            If InStr(CVal, vbLf) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbLf, "")
'            ElseIf InStr(CVal, vbCr) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbCr, "")
'            ElseIf InStr(CVal, vbCrLf) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbCrLf, "")
'            End If
            L_Ws.Cells(i, 2).Value = CVal
        Next i
        .Save
    End With

End Sub

Public Function Get_FildLis(ByVal str_RStn As String, str_LStn As String, str_Nullkey As String, Colnam As Long)
'★フィールドリストの作成
    'シートのフィールド行値を指定シート列へ縦転記
    '(引数1:取得したいシート範囲,引数2:転記したいリストシート名,引数:3Null除外フィールド名引数4:転記したい列番号)
    Dim L_Ws As Worksheet
    Dim i As Long

    Application.ScreenUpdating = False
    Set L_Ws = Sheets(str_LStn)
    L_Ws.Unprotect
    L_Ws.Columns(Colnam).Clear
    Call Opn_ExlRs(str_RStn, str_Nullkey)
    With Exl_Rs
        For i = 0 To .Fields.Count - 1
            L_Ws.Cells(i + 1, Colnam).Value = Exl_Rs.Fields(i).Name
        Next i
    End With
    Call Dis_Exl_Rs
    
End Function

Public Function Get_WHERELis(ByVal str_RStn As String, str_LStn As String, str_FCAd As String, Colnam As Long)
'★WHERE句リストの作成
    'シートのフィールド行値を指定シート列へ縦転記
    '(引数1:取得したいシート名,引数2:転記したいリストシート名,引数:3先頭フィールドセルアドレス,引数4:転記したい列番号)
    Dim L_Ws As Worksheet
    Dim R_Ws As Worksheet
    Dim i, sRow, sCol As Long

    Application.ScreenUpdating = False
    Set R_Ws = Sheets(str_RStn)
    Set L_Ws = Sheets(str_LStn)
    With R_Ws
        sRow = .Range(str_FCAd).Row
        sCol = .Range(str_FCAd).Column
        L_Ws.Columns(Colnam).Clear
        For i = sCol To 200
            L_Ws.Cells(i - sCol + 1, Colnam).Value = .Cells(sRow, i).Value
        Next i
    End With

End Function
