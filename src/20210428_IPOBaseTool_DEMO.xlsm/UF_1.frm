VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_1 
   Caption         =   "外部データカラムID登録フォーム"
   ClientHeight    =   7605
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8565
   OleObjectBlob   =   "UF_1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'★フォーム起動時リスト値読込
    Call Get_SearchD

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
    
    str_Skey = Me.TB_1.Value
    Call Get_SearchD(str_Skey)

End Sub

Private Sub CMD_2_Click()
'★登録ボタンクリック
    Dim eRow As Long
    Dim Ws As Worksheet
    Dim str_Ans As String
    
    Set Ws = Sheets("カラム設定")
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
    End With
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "管理表カラムIDを設定してから行ってください", 16, "管理表カラム未入力エラー"
            Exit Sub
        End If
        .Cells(eRow, 7).Value = str_Ans
    End With
    Unload UF_1


End Sub

Private Sub CMD_3_Click()
'★戻るボタンクリック
    Unload UF_1

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
    Dim Ws As Worksheet
    Dim str_Ans As String
    
    Set Ws = Sheets("カラム設定")
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
    End With
    With Ws
        eRow = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row
        If .Cells(eRow, 7).Value <> "" Then
            MsgBox "管理表カラムIDを設定してから行ってください", 16, "管理表カラム未入力エラー"
            Exit Sub
        End If
        .Cells(eRow, 7).Value = str_Ans
    End With
    Unload UF_1

End Sub

Public Function Get_SearchD(Optional ByVal str_Skey As Variant = "")
'★検索内容でレコードセット生成⇒リストボックス反映
    Const adOpenKeyset = 1, adLockReadOnly = 1
    Dim str_RCn  As String
    Dim R_Cn As ADODB.Connection
    Dim R_Rs As ADODB.Recordset
    Dim str_SQL As String
    Dim eRow As Long
    Dim R_Ws As Worksheet

    Set R_Ws = Sheets("T_GAIBColList")
    eRow = R_Ws.Cells(Rows.Count, 5).End(xlUp).Row
 '読出データセット *******************************************************************
    Set R_Cn = New ADODB.Connection
    Set R_Rs = New ADODB.Recordset
    If R_Cn.State = 1 Then End
    R_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    R_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=NO;IMEX=1"
    R_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT * "
    str_SQL = str_SQL & " FROM [T_GAIBColList$A3:B500] "
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



