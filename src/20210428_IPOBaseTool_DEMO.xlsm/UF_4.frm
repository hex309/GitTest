VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_4 
   Caption         =   "表示したい項目を選んでください"
   ClientHeight    =   3465
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12360
   OleObjectBlob   =   "UF_4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
'★起動時リスト読込
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
    Dim eCol As Long
    Dim str_Ans As String
    
     eCol = ActiveSheet.Range("B7").End(xlToRight).Column
    
    If Me.TB_2.Value = "" Then
        MsgBox "IDが選択されていません", 16
        Exit Sub
    End If
    str_Ans = Me.TB_2.Value
    ActiveSheet.Unprotect
    ActiveSheet.Cells(7, eCol + 1).Value = Me.TB_2.Value
    ActiveSheet.Range("G:HZ").EntireColumn.AutoFit
    Unload UF_0
    Call St_Lock

End Sub

Private Sub CMD_3_Click()
'★閉じるボタンクリック
    Unload UF_4

End Sub

Private Sub ListBox1_Click()
'★リストボックスクリックイベント
    With Me.ListBox1
        Me.TB_2.Value = .List(.ListIndex, 1)
    End With
    
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'★リストボックスダブルクリックイベント
    Dim eCol As Long
    Dim str_Ans As String
    
    eCol = ActiveSheet.Range("B7").End(xlToRight).Column
    With Me.ListBox1
        str_Ans = .List(.ListIndex, 1)
        ActiveSheet.Unprotect
        ActiveSheet.Cells(7, eCol + 1).Value = str_Ans
        ActiveSheet.Range("G:HZ").EntireColumn.AutoFit
    End With
    Unload UF_4
    Call St_Lock

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



