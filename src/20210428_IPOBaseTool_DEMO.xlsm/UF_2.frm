VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_2 
   Caption         =   "外部データカラムID登録フォーム"
   ClientHeight    =   7400
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   13680
   OleObjectBlob   =   "UF_2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'★フォーム起動時リスト値読込&新規登録ID読込
    
    Call Get_SearchD1
    Call Get_SearchD2
    Me.TB_2.Value = Sheets("管理表新規登録").Range("D6").Value
    Me.ListBox1.Clear
    Me.Repaint
    Me.TB_0.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'★Xで閉じられなくする
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
End Sub

Private Sub CMD_1_Click()
'★検索ボタンクリック
    Dim str_Skey As Variant
    
    str_Skey = Me.TB_0.Value
    Call Get_SearchD1(str_Skey)

End Sub

Private Sub CMD_2_Click()
'★検索ボタンクリック
    Dim str_Skey As Variant
    
    str_Skey = Me.TB_5.Value
    Call Get_SearchD2(str_Skey)

End Sub

Private Sub CMD_3_Click()
'★登録ボタンクリック
    Dim str_ID0 As String
    Dim str_ID1 As String
    Dim str_ID2 As String
    Dim Ans As Long
    
    With Me
        str_ID0 = .TB_2.Value
        str_ID1 = .TB_3.Value
        str_ID2 = .TB_4.Value
        If .TB_2.Value = "" Or .TB_3.Value = "" Or .TB_4.Value = "" Then
            MsgBox "紐付ID情報を全て設定してから行ってください", 16, "未入力エラー"
            Exit Sub
        End If
    End With
    Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID0 & "'")
    With Ac_Rs
        If IIf(IsNull(!T_2), "", !T_2) <> "" Then
            Ans = MsgBox("この管理表IDには既に紐付ID登録されていますが" & vbCrLf & _
                                    "表示中のIDで上書き登録しますか?" & vbCrLf & vbCrLf & _
                                    "はい　　で上書き" & vbCrLf & _
                                    "いいえ　でキャンセルします", _
                                    vbYesNo + vbInformation, "既に紐付IDが登録されています")
            If Ans = vbYes Then
                !T_2 = str_ID1
                !T_3 = str_ID2
                .Update
                MsgBox "紐付登録が完了しました", vbInformation
                Ans = MsgBox("登録結果を確認しますか？" & vbCrLf & _
                "はい 　 で管理表編集画面に" & vbCrLf & _
                "いいえ　でホーム画面に移動します", vbYesNo + vbInformation, "結果の確認")
                If Ans = vbYes Then
                    Call vis_KANRISt
                    With Sheets("管理表編集登録")
                        .Range("D4").Value = str_ID0
                        .Range("E4").Value = str_ID1
                        .Range("F4").Value = str_ID2
                    End With
                    Call Run_Search_Costumvew("管理表編集登録")
                    Call Re_Scrl
                ElseIf Ans = vbNo Then
                    Call vis_UISt
                End If
            ElseIf Ans = vbNo Then
                MsgBox "キャンセルされました", vbInformation
                Exit Sub
            End If
        Else
            !T_2 = str_ID1
            !T_3 = str_ID2
            .Update
            MsgBox "紐付登録が完了しました", vbInformation
            Ans = 0
            Ans = MsgBox("登録結果を確認しますか？" & vbCrLf & _
            "はい 　 で管理表編集画面に" & vbCrLf & _
            "いいえ　でホーム画面に移動します", vbYesNo + vbInformation, "結果の確認")
            If Ans = vbYes Then
                Call vis_KANRISt
                With Sheets("管理表編集登録")
                    .Unprotect
                    .Range("D4").Value = str_ID0
                    .Range("E4").Value = str_ID1
                    .Range("F4").Value = str_ID2
                End With
                Call Run_Search_Costumvew("管理表編集登録")
                Call Re_Scrl
            ElseIf Ans = vbNo Then
                Call vis_UISt
            End If
    End If

    End With
    Call Dis_Ac_Rs
    Unload UF_2

End Sub

Private Sub CMD_4_Click()
'★戻るボタンクリック
    Unload UF_2

End Sub

Private Sub ListBox1_Click()
'★リストボックスクリックイベント
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 0)
    End With
    
End Sub

Private Sub ListBox2_Click()
'★リストボックスクリックイベント
    With Me.ListBox2
        Me.TB_3.Value = .List(.ListIndex, 0)
        Me.TB_4.Value = .List(.ListIndex, 1)
    End With
    
End Sub

Public Function Get_SearchD1(Optional ByVal str_Skey As Variant = "")
'★検索内容でレコードセット生成⇒リストボックス反映　管理表ID用
    Dim str_SQL As String
 '読出データセット
     If str_Skey <> "" Then
        str_SQL = str_SQL & " AND T_1 LIKE'%" & str_Skey & "%'"
    End If
    Call Opn_AcRs("T_KANRI", "T_1", str_SQL)
 'リストボックスに追加
    With Me.ListBox1
        .Clear
        Do Until Ac_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(Ac_Rs!T_1), "", Ac_Rs!T_1)
            Ac_Rs.MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

End Function

Public Function Get_SearchD2(Optional ByVal str_Skey As Variant = "")
'★検索内容でレコードセット生成⇒リストボックス反映　外部データID用
    Dim str_SQL As String
 '読出データセット
     If str_Skey <> "" Then
        str_SQL = str_SQL & " AND F_1 LIKE'%" & str_Skey & "%'"
    End If
    Call Opn_AcRs("T_GAIBU1", "F_1", str_SQL)
 'リストボックスに追加
    With Me.ListBox2
        .Clear
        Do Until Ac_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(Ac_Rs!F_1), "", Ac_Rs!F_1)
            .List(.ListCount - 1, 1) = IIf(IsNull(Ac_Rs!F_2), "", Ac_Rs!F_2)
            Ac_Rs.MoveNext
        Loop
    End With
    Call Dis_Ac_Rs

End Function
