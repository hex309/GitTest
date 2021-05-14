VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProductItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'// ************************************************************************************************************************
'//  モジュール名　　：CollectionItems
'//  概要　　　　　　：Itemsクラスのサンプル
'//  参照設定　　　　：なし
'//  作成者　　　　　：Nakabayashi Toshiharu
'//  作成日　　　　　：2013/05/24
'//  更新日：内容　　：
'// ************************************************************************************************************************

Private mItems As Collection
Private mItemDictionary As Object

Private Sub Class_Initialize()
    Set mItems = New Collection
    Set mItemDictionary = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set mItems = Nothing
    Set mItemDictionary = Nothing
End Sub

'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：Add
'//  機能　　　　　　：vData：データ
'//  引数　　　　　　：
'//  戻り値　　　　　：なし
'//  作成者　　　　　：Nakabayashi Toshiharu
'//  作成日　　　　　：2013/05/24
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Public Sub Add(ByVal vData As String, ByVal vNum As Long _
    , ByVal vSupplier As String)
    Dim temp As ProductItem

    On Error Resume Next
    mItemDictionary.Add KEY:=vData, Item:=mItems.Count + 1

    If Err.Number = 0 Then
        Set temp = New ProductItem
        mItems.Add temp
    Else
        Set temp = mItems.Item(ExistsData(vData))
    End If

    With temp
        .ProductID = vData
        .Num = vNum
        .Supplier = vSupplier
    End With

    Set temp = Nothing
End Sub

Public Function GetAllData() As Variant
    Dim temp() As Variant
    Dim i As Long
    For i = 1 To Me.Count
        ReDim Preserve temp(1 To 3, 1 To i)
        temp(1, i) = Me.Item(i).ProductID
        temp(2, i) = Me.Item(i).Num
        temp(3, i) = Me.Item(i).Supplier
    Next
    GetAllData = Application.WorksheetFunction.Transpose(temp)
End Function
'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：ExistsData
'//  機能　　　　　　：存在チェック
'//  引数　　　　　　：vData：検索データ
'//  戻り値　　　　　：インデックス番号
'//  作成者　　　　　：Nakabayashi Toshiharu
'//  作成日　　　　　：2013/05/24
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Public Function ExistsData(ByVal vData As String) As Long
    ExistsData = 0

    If mItemDictionary.Exists(vData) Then
        ExistsData = mItemDictionary.Item(vData)
    End If
End Function

Public Property Get Item(ByVal Index As Long) As ProductItem
    Set Item = mItems.Item(Index)
End Property

Public Property Get Count()
    Count = mItems.Count
End Property

