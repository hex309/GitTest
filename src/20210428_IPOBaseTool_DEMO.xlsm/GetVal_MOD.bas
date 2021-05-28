Attribute VB_Name = "GetVal_MOD"
Option Explicit
'★値取得系モジュール

Public Function Get_Maxval() As String
'★ID最大値取得
    Dim str_Ans As Variant
    Dim cnt As Long
    
    Sheets("CHK_MID").Range("A2:A80000").ClearContents
    Call Opn_AcRs("T_KANRI", "T_1", , "T_1")
    With Ac_Rs
        cnt = 1
        Do Until .EOF
            str_Ans = ""
            str_Ans = !T_1
            str_Ans = Mid(str_Ans, 4, Len(str_Ans))
            cnt = cnt + 1
            Sheets("CHK_MID").Cells(cnt, 1).Value = str_Ans
            .MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

    str_Ans = ""
    str_Ans = Sheets("CHK_MID").Range("B1").Value
    Debug.Print str_Ans
    Get_Maxval = "XXX" & str_Ans + 1

End Function

Public Function Get_ChangeData()
'★更新処理完了時メッセージデータ詳細部作成
    '管理表上書き更新処理前に使用
    Dim str_SKey1, str_SKey2, str_SKey3, str_Ans As String
    
    Call Opn_ExlRs("管理表編集登録$B7:FZ8000", "T_1", " AND RegFlg='有'")
    If Exl_Rs.EOF = True Then ''有'データがなかった場合
        MsgBox "変更されたデータはありません", vbInformation
        Call Dis_Exl_Rs
        End
    End If
    Do Until Exl_Rs.EOF '読出しデータから更新メッセージ作成
        str_SKey1 = str_SKey1 & Exl_Rs!T_1 & vbCrLf
        str_SKey2 = str_SKey2 & Exl_Rs!T_2 & "," & Exl_Rs!T_3 & vbCrLf
        Exl_Rs.MoveNext
    Loop
    str_Ans = str_Ans & "更新されたレコードは" & vbCrLf & "【管理表キー】" & vbCrLf
    str_Ans = str_Ans & str_SKey1 & "【外部データ２キー】" & vbCrLf
    str_Ans = str_Ans & str_SKey2 & "でした"
    Call Dis_Exl_Rs
    Get_ChangeData = str_Ans

End Function

Public Function Get_FilFol(ByVal Flg As Long, Optional F_type As String = "") As Variant
'★ダイアログからファイル/フォルダパス取得ファンクション
    '(引数1:ファイル/フォルダ選択フラグ 1=ファイル　2=フォルダ,引数2:メッセージとファイル拡張子）
    'F_type=サンプル:"インポートファイルを選択してください (*.xlsb;*.xlsx;*.accdb), *.xlsb;*.xlsx;*.accdb"
    Dim Sfile As String
    Dim i As Integer
    Dim s As String
   
    If Flg = 1 Then
        Sfile = Application.GetOpenFilename(F_type)
        If Sfile = "False" Then
            Get_FilFol = ""
            Exit Function
        End If
            Get_FilFol = Sfile
    ElseIf Flg = 2 Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Show
            If Sfile = "False" Then
                Get_FilFol = ""
            Else
                Get_FilFol = .SelectedItems(1)
                
            End If
        End With
    End If

End Function
