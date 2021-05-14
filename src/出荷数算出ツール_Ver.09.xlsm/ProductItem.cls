VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// ************************************************************************************************************************
'//  モジュール名　　：CollectionItem
'//  概要　　　　　　：Itemクラスのサンプル
'//  参照設定　　　　：なし
'//  作成者　　　　　：Nakabayashi Toshiharu
'//  作成日　　　　　：2013/05/24
'//  更新日：内容　　：
'// ************************************************************************************************************************

Private mProductID As String   '品目
Private mSupplier As String
Private mNum As Long            '数量

Public Property Get ProductID() As String
    ProductID = mProductID
End Property
Public Property Let ProductID(ByVal vData As String)
    mProductID = vData
End Property

Public Property Get Supplier() As String
    Supplier = mSupplier
End Property
Public Property Let Supplier(ByVal vData As String)
    mSupplier = vData
End Property

Public Property Get Num() As Long
    Num = mNum
End Property
Public Property Let Num(ByVal vData As Long)
    mNum = mNum + vData
End Property

