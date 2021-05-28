Attribute VB_Name = "CHK_MOD"
Option Explicit
'★チェック系モジュール
Public Function CHK_Duplicate_ID(ByVal str_Tbln As String, _
                         str_Sval As String, str_Feldn As String) As Boolean
'★値重複チェック
    Call Opn_AcRs(str_Tbln, str_Feldn, , str_Feldn)
    With Ac_Rs
        Do Until .EOF
            If str_Sval = Ac_Rs(str_Feldn).Value Then GoTo Skip
            .MoveNext
        Loop
    End With
    Call Dis_Ac_Rs
    CHK_Duplicate_ID = False
    Debug.Print CHK_Duplicate_ID
    Exit Function
Skip:
    CHK_Duplicate_ID = True
    Debug.Print CHK_Duplicate_ID
    Call Dis_Ac_Rs
    
End Function
Sub ttttteeetttst1()
    Debug.Print CHK_Duplicate_ID("T_KANRI", "XXX190201003", "T_1")
End Sub

Public Sub CHK_RegChange(Optional ByVal str_Stn As String = "管理表編集登録")
'★変更データ有無チェック　管理表編集_登録シート⇔T_KANRI
    'DB比較で変更されたものにシートカラム更新有無='有' 管理表上書更新処理前に使用
    Dim i, r, eRow As Long
    Dim str_ID, str_Fildn As String
    Dim AcVal, ExlVal As Variant
    Dim str_RngAd As String
    
    With Sheets(str_Stn)
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        If eRow > 40 Then 'チェック量が多いと不安定になる為(50くらいが限界ぽい）
            MsgBox "レコード数が多すぎます" & vbCrLf & "レコード数を30件以内に絞込してください", 16
            End
        End If
        For i = 10 To eRow
            str_ID = .Cells(i, 4).Value
            Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID & "'")
            str_RngAd = Sheets(str_Stn).Range("B7").End(xlToRight).Address
            str_RngAd = Replace(str_RngAd, "7", "50")
            str_RngAd = Replace(str_RngAd, "$", "")
            Call Opn_ExlRs(str_Stn & "$D7:" & str_RngAd, "T_1", " AND T_1='" & str_ID & "'")
            For r = 3 To Exl_Rs.Fields.Count - 1
                str_Fildn = Exl_Rs.Fields(r).Name
                AcVal = IIf(IsNull(Ac_Rs(str_Fildn).Value), "", Ac_Rs(str_Fildn).Value)
                ExlVal = IIf(IsNull(Exl_Rs(str_Fildn).Value), "", Exl_Rs(str_Fildn).Value)
                If AcVal <> ExlVal Then
                    .Unprotect
                    .Cells(i, 2) = "有"
                    GoTo Skip
                End If
            Next r
Skip:
            Call Dis_Exl_Rs
            Call Dis_Ac_Rs
        Next i
    End With
        
End Sub

Public Function CHK_WFildsNam(ByVal str_Stn As String, Coln As Long, _
                                                    sRow As Long, chk_Val As String) As Boolean
'★シートの値重複チェック
'(引数1:シート名,引数2:カラム列No,引数3:先頭行No,引数4:検索値)
    Dim eRow, i As Long
    Dim Ws As Worksheet
    
    Set Ws = Sheets(str_Stn)
    With Ws
        eRow = .Cells(Rows.Count, Coln).End(xlUp).Row
        For i = sRow To eRow
            If .Cells(i, Coln) = chk_Val Then GoTo Skip
        Next i
    End With
    CHK_WFildsNam = False
    Exit Function
Skip:
    CHK_WFildsNam = True
 
End Function
