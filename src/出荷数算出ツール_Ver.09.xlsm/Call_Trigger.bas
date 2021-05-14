Attribute VB_Name = "Call_Trigger"
Option Explicit

'============================================================
'   ユーザが呼べるマクロ
'============================================================

Public Sub 検収_入荷実績一覧ファイル選択()
    Call SubSetCSVFilePathByUserChoice(uexlOrderArrival)
End Sub

Public Sub 商品台帳_前月_ファイル選択()
    Call SubSetCSVFilePathByUserChoice(uexlPreProductBook)
End Sub

Public Sub 商品台帳_当月_ファイル選択()
    Call SubSetCSVFilePathByUserChoice(uexlProductBook)
End Sub

Public Sub 保存先フォルダ選択()
    Call GetFoldePath
End Sub

Public Sub 出荷数算出()
    Call GetCurrentData
End Sub

Public Sub 仕入先_ファイル選択()
    Call SubSetCSVFilePathByUserChoice(uexlSupplierBook)
End Sub

Public Sub 仕入先一覧取得()
    Call GetSupplier
End Sub

