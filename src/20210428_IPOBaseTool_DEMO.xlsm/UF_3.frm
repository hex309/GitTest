VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_3 
   Caption         =   "外部データカラムID登録フォーム"
   ClientHeight    =   4515
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11040
   OleObjectBlob   =   "UF_3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub UserForm_Activate()
'★フォーム起動時リスト値読込&新規登録ID読込
    Call Get_SearchD1
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

Private Sub CMD_3_Click()
'★削除ボタンクリック
    Dim str_ID0 As String
    Dim Ans As Long
    
    With Me
        str_ID0 = .TB_2.Value
        If str_ID0 = "" Then
            MsgBox "IDが入力されてません", 16, "未入力エラー"
            Exit Sub
        End If
    End With
    Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID0 & "'")
    With Ac_Rs
        Set Ac_Cmd = New ADODB.Command
        str_SQL = ""
        str_SQL = str_SQL & "DELETE FROM T_KANRI"
        str_SQL = str_SQL & " WHERE T_1='" & str_ID0 & "'"
        With Ac_Cmd
            .ActiveConnection = Ac_Cn
            .CommandText = str_SQL
            .Execute
        End With
        Ans = MsgBox("このIDのデータが管理表DBから削除されます" & vbCrLf & _
                                "削除しますか?" & vbCrLf & vbCrLf & _
                                "はい　　で削除" & vbCrLf & _
                                "いいえ　でキャンセルします", _
                                vbYesNo + vbInformation, "削除確認")
        If Ans = vbYes Then
            MsgBox "削除が完了しました", vbInformation
        ElseIf Ans = vbNo Then
            MsgBox "キャンセルされました", vbInformation
            Exit Sub
        End If
    End With
    Call Dis_Ac_Rs
    Unload UF_3

End Sub

Private Sub CMD_4_Click()
'★戻るボタンクリック
    Unload UF_3

End Sub

Private Sub ListBox1_Click()
'★リストボックスクリックイベント
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 0)
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
