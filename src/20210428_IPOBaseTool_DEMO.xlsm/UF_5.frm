VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_5 
   Caption         =   "お気に入り保存"
   ClientHeight    =   4250
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5775
   OleObjectBlob   =   "UF_5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
    End If
    
End Sub

Private Sub CMD_2_Click()
'★登録ボタンクリック
    Dim eRow As Long
    Dim str_Ans As String

    str_Ans = Me.TB_2.Value
    If str_Ans = "" Then
        MsgBox "登録名が入力されていません", 16
        Exit Sub
    End If
    With Sheets("カスタム編集登録お気に入り")
        eRow = .Cells(Rows.Count, 1).End(xlUp).Row + 1
        Call Opn_ExlRs("カスタム編集登録お気に入り$A1:B1000", "登録名")
        With Exl_Rs
            Do Until Exl_Rs.EOF
               If !登録名 = str_Ans Then GoTo Skip
                .MoveNext
            Loop
        End With
        Call Dis_Exl_Rs
        .Unprotect
        .Cells(eRow, 1).Value = str_Ans
        Sheets("管理表編集登録").Range("G7:GU7").Copy
        .Cells(eRow, 2).PasteSpecial Paste:=xlValues
    End With
    ThisWorkbook.Save
    MsgBox "登録完了！", vbInformation
    Unload UF_5
    Call St_Lock
    
    Exit Sub
Skip:
    MsgBox "その名前は既に使われています", 16
    Call Dis_Exl_Rs
    Exit Sub

End Sub

Private Sub CMD_3_Click()
'★閉じるボタンクリック
    Unload UF_5

End Sub
