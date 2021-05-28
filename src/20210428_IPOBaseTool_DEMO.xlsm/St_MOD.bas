Attribute VB_Name = "St_MOD"
Option Explicit
'★シート操作系モジュール

Public Sub Run_FeldsSET_Seve()
'★フィールド設定保存反映実行
    Dim Ws As Worksheet
    
    Application.ScreenUpdating = False
    Set Ws = Sheets("管理表フィールド設定")
    Call Get_KANRIColList '管理表フィールド名リスト作成
    ThisWorkbook.Save
    MsgBox "設定の保存・反映が完了しました", vbInformation
    
End Sub

Sub Run_Col_Seve()
'★カラム設定保存反映実行
    Call vis_CosKANRISt
'    Call Col_Seve("TG_T_ColList", 0)
'    Call Col_Seve("TG_G_ColList", 1)
'    Call St_ColLis("TG_G_ColList", "外部データ", "B5")
    
    ThisWorkbook.Save
    
End Sub

Public Function Col_Seve(ByVal str_L_Stn As String, Flg As Long)
'★カラム設定保存反映
    '設定シートから設定カラム情報のみカラム情報シートに転記
    '引数1:書込シート名,保存カラム選択フラグ　０＝当社/1=外部
    Dim eRow As Long
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim str_StRng As String
  
    Set R_Ws = Sheets("カラム設定")
    Set L_Ws = Sheets(str_L_Stn)
 '読出データセット *******************************************************************
    If Flg = 0 Then '当社カラムデータ読出
        eRow = R_Ws.Cells(Rows.Count, 5).End(xlUp).Row
        str_StRng = "カラム設定$E4:E" & eRow
        Call Opn_ExlRs(str_StRng, "管理表カラムID")
    End If
    If Flg = 1 Then  '外部カラムデータ読出
        eRow = R_Ws.Cells(Rows.Count, 7).End(xlUp).Row
        str_StRng = "カラム設定$G4:G" & eRow
        Call Opn_ExlRs(str_StRng, "外部カラムID")
    End If
 '読出データセットここまで **************************************************************
 'データ転記
    With L_Ws
       .Unprotect
       .Cells.ClearContents
       .Range("A1").CopyFromRecordset Exl_Rs
    End With
    Call Dis_Exl_Rs

End Function

Public Function St_ColLis(ByVal str_R_Stn As String, str_L_Stn As String, str_rng As String)
'★外部データシートへの外部カラム設定反映
    '引数1:読出シート名,引数2:書込シート名,引数3:フィールド先頭セルアドレス
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim LC_Ws As Worksheet
    Dim str_LC_Stn As String
    Dim sRow, sCol As Long
    Dim cnt As Long

    Set R_Ws = Sheets(str_R_Stn)
    Set L_Ws = Sheets(str_L_Stn)

    sRow = L_Ws.Range(str_rng).Row
    sCol = L_Ws.Range(str_rng).Column
    
    Call Opn_ExlRs("カラム設定$G4:G300", "外部カラムID")

    With L_Ws
        .Unprotect
        .Range(.Cells(sRow, sCol), .Cells(sRow, 300)).ClearContents '既存フィールド名クリア
        cnt = 0
        Do Until Exl_Rs.EOF
            cnt = cnt + 1
            .Cells(sRow, sCol - 1 + cnt).Value = Exl_Rs!外部カラムID
            Exl_Rs.MoveNext
        Loop
    End With
    L_Ws.Range("G:GZ").EntireColumn.AutoFit
    Call Dis_Exl_Rs

End Function

Public Sub Re_KANRI()
'★管理表編集_登録シートリセット
    Dim K_Ws As Worksheet
    
    Set K_Ws = Sheets("管理表編集登録")
     With K_Ws
        .Unprotect
        .Rows(4).ClearContents
        .Rows(10).ClearContents
        .Range("11:100000").Delete
'        .Range("B10").Select
    End With
    Call Re_Scrl
    Call St_Lock

End Sub

Public Sub Re_CosKANRI()
'★管理表編集_登録シートリセット
    Dim K_Ws As Worksheet
    
    Set K_Ws = Sheets("管理表編集登録")
     With K_Ws
        .Unprotect
        .Rows(4).ClearContents
        .Rows(10).ClearContents
        .Range("11:100000").Delete
'        .Range("B10").Select
    End With
    Call Re_Scrl
    Call St_Lock

End Sub
Public Sub Run_Clear_SearchKey1()
'★検索条件クリア　管理表編集_登録用
    Dim Ws As Worksheet
    
    Set Ws = Sheets("管理表編集登録")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Re_KANRI

End Sub

Public Sub Run_Clear_SearchKey2()
'★検索条件クリア　管理表出力ビュー用
    Dim Ws As Worksheet
    
    Application.ScreenUpdating = False
    Set Ws = Sheets("管理表出力ビュー")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Run_Douki_KANRIvew

End Sub

Public Sub Run_Clear_SearchKey3()
'★検索条件クリア　 カスタムビュー用
    Dim Ws As Worksheet
    
    Set Ws = Sheets("カスタムビュー")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
    End With
    Call Run_Search_Costumvew

End Sub

Public Sub Run_Clear_SearchKey4()
'★検索条件クリア　 カスタム編集登録用
    Dim Ws As Worksheet
    
    Set Ws = Sheets("管理表編集登録")
    With Ws
        .Unprotect
        .Rows(4).ClearContents
        .Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = ""
    End With
    Call Re_CostumSt("管理表編集登録")
'    Call Run_Search_Costumvew("管理表編集登録")

End Sub
Public Sub Re_CostumSt(Optional ByVal str_Stn As String = "カスタムビュー")
'★設定前カスタムシートのデータリセット
    Dim L_Ws As Worksheet
    
    Set L_Ws = Sheets(str_Stn)
    With L_Ws
        .Unprotect
         .Range("11:20000").Delete
        .Range("B10:GZ20000").ClearContents
        .Range("G:HZ").EntireColumn.AutoFit '16777164
        If str_Stn = "カスタムビュー" Then
            .Range("11:1000").Interior.Color = 13434828
        ElseIf str_Stn = "管理表編集登録" Then
            .Range("11:1000").Interior.Color = 16777164
        End If
    End With
    Call St_Lock

End Sub

Public Sub Re_Costum_ALLclear()
'★カスタム設定のクリアリセット
    Dim Ans As Long
    
    Ans = MsgBox("表示中のレコードおよび" & vbCrLf & _
                            "先頭５列以外の設定カラムが全てクリアされます" & vbCrLf & _
                                "よろしいですか?", vbYesNo + vbInformation, "ファイナルアンサー")
    If Ans = vbNo Then
        End
    End If
    ActiveSheet.Unprotect
    ActiveSheet.Range("G5:GS5").ClearContents
    ActiveSheet.Range("G7:GS7").ClearContents
    ActiveSheet.Range("E10:GS80000").ClearContents
    ActiveSheet.Range("B10").Select
    Call Re_CostumSt("管理表編集登録")
    Call St_Lock

End Sub

Public Sub Clear_SearchKey_G()
'★外部データ検索条件クリア
    ActiveSheet.Unprotect
    Sheets("外部データ").Rows(3).ClearContents
    Call Run_Douki_GAIB
    Call St_Lock
    Call Re_Scrl

End Sub
