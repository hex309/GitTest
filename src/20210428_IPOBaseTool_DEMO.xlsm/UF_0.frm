VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_0 
   Caption         =   "管理表カラムID登録フォーム"
   ClientHeight    =   8010
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8565
   OleObjectBlob   =   "UF_0.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Activate()

    Call Get_SearchD

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
End Sub

Private Sub CMD_1_Click()
'★検索ボタンクリック
    Dim str_Skey As Variant
    
    str_Skey = Me.TB_1.Value
    
    Call Get_SearchD(str_Skey)

End Sub

Private Sub CMD_2_Click()
'★登録ボタンクリック
    Dim eRow As Long
    Dim str_Ans As String
    
    eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
    If Cells(eRow, 7).Value = "" Then
        MsgBox "管理表カラムIDに対応する外部カラムIDが設定されていません " & vbCrLf & _
        "まだ設定されていない外部カラムIDを設定してください", 16
        End
    End If
    If Me.TB_2.Value = "" Then
        MsgBox "IDが選択されていません", 16
        Exit Sub
    End If
    str_Ans = Me.TB_2.Value
    If CHK_WFildsNam("カラム設定", 6, 10, str_Ans) = True Then
            MsgBox "ID：" & str_Ans & "　は登録済みです" & vbCrLf & _
            "管理表カラム登録できるIDは1つだけです " & vbCrLf & _
            "既に登録したIDの登録はできません", 16, "重複エラー！"
            Exit Sub
        End If
    ActiveSheet.Cells(eRow + 1, 5).Value = Me.TB_2.Value
    Unload UF_0

End Sub

Private Sub CMD_3_Click()
'★戻るボタンクリック
    Unload UF_0

End Sub

Private Sub ListBox1_Click()
'★リストボックスクリックイベント
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 1)
    End With
    
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'★リストボックスダブルクリックイベント
    Dim eRow As Long
    Dim str_Ans As String
    
    eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
    If Cells(eRow, 7).Value = "" Then
        MsgBox "管理表カラムIDに対応する外部カラムIDが設定されていません " & vbCrLf & _
        "まだ設定されていない外部カラムIDを設定してください", 16
        End
    End If
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
        If CHK_WFildsNam("カラム設定", 6, 10, str_Ans) = True Then
            MsgBox "ID：" & str_Ans & "　は登録済みです" & vbCrLf & _
            "管理表カラム登録できるIDは1つだけです " & vbCrLf & _
            "既に登録したIDの登録はできません", 16, "重複エラー！"
            Exit Sub
        End If
         ActiveSheet.Cells(eRow + 1, 5).Value = str_Ans
    End With
    Unload UF_0

End Sub

Public Function Get_SearchD(Optional ByVal str_Skey As Variant = "")
'★検索内容でレコードセット生成⇒リストボックス反映
    Const adOpenKeyset = 1, adLockReadOnly = 1
    Dim str_RCn  As String
    Dim R_Cn As ADODB.Connection
    Dim R_Rs As ADODB.Recordset
    Dim str_SQL As String
 '読出データセット *******************************************************************
    Set R_Cn = New ADODB.Connection
    Set R_Rs = New ADODB.Recordset
    If R_Cn.State = 1 Then End
    R_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    R_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=NO;IMEX=1"
    R_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT * "
    str_SQL = str_SQL & " FROM [T_KANRIColList$A6:B500] "
    If str_Skey <> "" Then
        str_SQL = str_SQL & " WHERE F2 LIKE'%" & str_Skey & "%'"
    End If
    R_Rs.Open str_SQL, R_Cn, adOpenKeyset, adLockReadOnly
 '読出データセットここまで **************************************************************
 'リストボックスに追加
    With Me.ListBox1
        .Clear
        Do Until R_Rs.EOF
            .AddItem ""
            .List(.ListCount - 1, 0) = IIf(IsNull(R_Rs!F2), "", R_Rs!F2)
            .List(.ListCount - 1, 1) = R_Rs!F1
            R_Rs.MoveNext
        Loop
    End With
'◆後処理
    R_Rs.Close 'レコードセットのクローズ
    Set R_Rs = Nothing
    R_Cn.Close 'コネクションのクローズ
    Set R_Cn = Nothing  'オブジェクトの破棄

End Function



