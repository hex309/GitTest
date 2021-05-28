Attribute VB_Name = "Get_SQL_MOD"
Option Explicit
'★SQL文に必要な文字を取得成形モジュール

Public Function Get_WHERE(ByVal str_Stn As String, str_FC As String, str_VC As String) As String
'★SQL追加WHERE句の作成
    '(引数1:取得したいシート名,引数２:検索フィールドの先頭行セルアドレス,引数3:検索値の先頭行セルアドレス)
    Dim i As Long
    Dim str_Ans As String
    Dim rs As ADODB.Recordset
    
    Application.ScreenUpdating = False
    Call Get_WHERELis(str_Stn, "T_WHEREList", str_FC, 1)
    Call Get_WHERELis(str_Stn, "T_WHEREList", str_VC, 2)
    Call Opn_ExlRs("T_WHEREList$A1:B200", "F1", , , 1)
    With Exl_Rs
        Do Until .EOF
            str_Ans = str_Ans & " AND " & !F1 & " Like '%" & !F2 & "%'"
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Get_WHERE = str_Ans
    
End Function

Public Function Get_SQLFelds(ByVal str_RStn As String) As String
'★SQL文フィールド指定部文取得
    'フィールドからカラム指定部文を生成
    Dim R_Ws As Worksheet
    Dim i, eCol, eRow As Long
    Dim str_Ans As String
    
    Set R_Ws = Sheets(str_RStn)
    With R_Ws
        str_Ans = ""
        eRow = .Cells(Rows.Count, 1).End(xlUp).Row
        For i = 1 To eRow
            str_Ans = str_Ans & .Cells(i, 1).Value & ","
        Next i
    End With
    str_Ans = Left(str_Ans, Len(str_Ans) - 1)
    Get_SQLFelds = str_Ans

End Function
