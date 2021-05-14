Attribute VB_Name = "▲Public宣言▲"
Option Explicit

'----------------------------------
'スリープ及びメッセージボックス（API宣言）
'----------------------------------
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "KERNEL32" (ByVal ms As LongPtr)
    Public Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#Else
    Public Declare Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
    Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#End If

Public Const MB_TOPMOST = 262144 ' &H40000
Public Const MB_OK = 0 ' &H0
Public Const MB_ICONINFOMATION = 64
Public Const MB_EXCLAMATION = 48


'----------------------------------
' フォーム画面最前面（API宣言）
'----------------------------------
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

'----------------------------------------------------
'パブリッククラス
'----------------------------------------------------
Public PubClsLucasAuth As LucasAuth:

'----------------------------------
'停止フラグ(StopButtonフォーム)
'----------------------------------
Public fStop As Boolean

Public Const IS_TEST As Boolean = False
'----------------------------------------------------
'パブリックIEオブジェクト
'----------------------------------------------------
Public oPubIE1  As InternetExplorerMedium

Public Pubシート名 As String
Public Pub操作番号 As Long

'----------------------------------------------------
'オートパイロット準備変数（見積登録編）
'----------------------------------------------------
Public Pubオートパイロット番号 As String
Public Pub制御シート名 As String
Public Pub見積シート名 As String

Public Pub見積登録件名 As String
Public Pub営業者コード As String
Public Pub主任者コード As String

Public Pub工期FROM As Variant
Public Pub工期TO As Variant

Public Pub店舗コード As String     '共通_工事発注仕様書用
Public Pub工事発注仕様書 As String '共通_工事発注仕様書用

Public Pub見積前提条件 As String '共通_見積登録アップロード用

'----------------------------------------------------
'シート名
'----------------------------------------------------
Public Const WSNAME_TOOL As String = "品目記入型 (ツール)"
Public Const WSNAME_VAL_HINMOKU As String = "変数一覧_品目管理"
Public Const WSNAME_VAL_MODULE As String = "変数一覧_モジュール"
Public Const WSNAME_LOG As String = "実行履歴"
Public Const WSNAME_LOGIN As String = "ログイン"
Public Const WSNAME_CSVUP As String = "CSVアップロート"
Public Const WSNAME_WARIKOMI As String = "割込T01"
Public Const WSNAME_CONFIG As String = "設定"
Public Const WSNAME_HINMOKU_MODULE As String = "品目記入型（モジュール）"
Public Const WSNAME_HINMOKU As String = "品目管理表"
Public Const WSNAME_CODE As String = "コードリスト"
Public Const WSNAME_FORMAT As String = "見積確認フォーマット"

'----------------------------------------------------
'入力ファイル　マスターファイル関連（回線申請一覧ファイル）
'----------------------------------------------------
Public Pubマスタブックフルパス As Variant 'C:\xxx\yyy\zzz.xlsx
Public Pubマスタブックパス As String      'C:\xxx\yyy\
Public Pubマスタブック名 As String        'zzz.xlsx

'----------------------------------------------------
'出力ファイル　作業手順書ファイル関連
'----------------------------------------------------
Public Pubブックフルパス As String 'C:\xxx\yyy\zzz.xlsx
'Public Pubブックパス As String      'C:\xxx\yyy\ 'Pub作業フォルダパス変数があるから未使用
Public Pubブック名 As String        'zzz.xlsx
Public Pub拡張子 As String          'xlsx

'----------------------------------------------------
'複製ファイル　マスターファイルの複製品（回線申請一覧ファイル）
'----------------------------------------------------
'Public Pub複製ブックフルパス As Variant 'C:\xxx\yyy\zzz.xlsx
'Public Pub複製ブックパス As String      'C:\xxx\yyy\
'Public Pub複製ブック名 As String        'zzz.xlsx

'----------------------------------------------------
'日時 出力ファイル名に日時付与で使用
'----------------------------------------------------
Public YYYYMMDD_HHNNSS '年月日_時分秒
Public YYYYMMDD        '年月日
Public HHNNSS          '時分秒

'----------------------------------------------------
'出力ファイルの作業用フォルダパス
'----------------------------------------------------
Public Pub作業フォルダパス As String 'C:\xxx\yyy\

'----------------------------------------------------
'新規レコードフラグ
'----------------------------------------------------
'Public Pub新規拠点名 As Boolean
'Public Pub新規商品名 As Boolean

Public Pub本文アドレス As String
