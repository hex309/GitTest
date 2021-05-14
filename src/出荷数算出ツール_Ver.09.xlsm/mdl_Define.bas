Attribute VB_Name = "mdl_Define"
Option Explicit
Option Private Module

'============================================================
'   グローバル定数・汎用列挙型
'============================================================


'----------------
'　定数
'----------------
'ファイル設定（兼シート名）
Public Const G_FILE_TARGET_OA   As String = "検収・入荷実績一覧"
Public Const G_FILE_TARGET_SS   As String = "商品台帳（期末）"
Public Const G_FILE_TARGET_DD   As String = "商品台帳（期首）"
Public Const G_FILE_TARGET_YA   As String = "出荷実績"
Public Const G_FILE_TARGET_SU   As String = "仕入先一覧"


'項目リストライン（各ファイル内項目名）←貼り付け先の名称に合わせないで下さい
Public Const G_ITEMLINE_OA      As String = "出荷予定,部品,数量,区分" 'X_受注引当側項目名
Public Const G_ITEMLINE_SS      As String = "行ラベル,設定数" '安全在庫数側項目名


'ツール起動シート
Public Const G_SHEETNAME_TOOL       As String = "元データ読込"

'計算用シート
Public Const G_SHEETNAME_CALCBASE   As String = "計算用"

'メインシート
Public Const G_SHEETNAME_MAIN       As String = "算出ツール"


'算出用ピボット名
Public Const G_PIVOTNAME_MAIN       As String = "pivot在庫"


'メインシート内展開ターゲット
Public Const G_PAST_TARGET_PARTNUMBER   As String = "部品番号"
Public Const G_PAST_TARGET_CALCRESULT   As String = "現在在庫数"


'安全在庫数ファイル内の安全在庫の列位置
Public Const C_安全在庫数列         As Long = 3



'拡張子
'取込用ＣＳＶファイルの拡張子
Public Const G_EXT_CSV          As String = ".csv"

'テキストファイルの拡張子
Public Const G_EXT_TXT          As String = ".txt"

'Excelファイルの拡張子
Public Const G_EXT_XLSX         As String = ".xlsx"
Public Const G_EXT_XLS          As String = ".xls*"

'ファイルフィルター用文字列
Public Const G_FILTERNAME_ALL   As String = "すべてのファイル(*.*)"
Public Const G_FILTERNAME_XLSX  As String = "Excelブック(*.xlsx)"
Public Const G_FILTERNAME_EXCEL As String = "Excelファイル(*.xls*)"
Public Const G_FILTERNAME_CSV   As String = "CSV形式(*.csv)"

'文字列フォーマット形式
Public Const G_FORMAT_DATE_YYYYMMDD         As String = "yyyymmdd"


'エラーハンドリング用設定
Public Const G_CTRL_ERROR_NUMBER_USER_NOTICE    As Long = vbObjectError + 1
Public Const G_CTRL_ERROR_NUMBER_USER_CAUTION   As Long = vbObjectError + 9
Public Const G_CTRL_ERROR_NUMBER_DEVELOPER      As Long = vbObjectError + 100



'----------------
'　列挙型
'----------------
'ファイルの種類
Public Enum ueXLFileType
    [_MIN] = 0
    uexlOrderArrival = 1                         '検収_入荷実績
    uexlPreProductBook = 2                       '商品台帳（前月）
    uexlProductBook = 3                          '商品台帳（当月）
    uexlSupplierBook = 4                         '仕入先
    [_MAX]
End Enum

Public Enum ueRC
    [_MIN] = 0
    uercRow = 1
    uercCol = 2
    [_MAX]
End Enum

'列番号の数値化
Public Enum ueColumnNum
    ueColA = 1
    ueColB = 2
    ueColC = 3
    ueColD = 4
    ueColE = 5
    ueColF = 6
    ueColG = 7
    ueColH = 8
    ueColI = 9
    ueColJ
    ueColK
    ueColL
    ueColM
    ueColN
End Enum

'上記を定数化
Public Const COL_A     As Long = ueColA
Public Const COL_B     As Long = ueColB
Public Const COL_C     As Long = ueColC
Public Const COL_D     As Long = ueColD
Public Const COL_E     As Long = ueColE
Public Const COL_F     As Long = ueColF
Public Const COL_G     As Long = ueColG
Public Const COL_H     As Long = ueColH
Public Const COL_I     As Long = ueColI
Public Const COL_J     As Long = ueColJ
Public Const COL_K     As Long = ueColK
Public Const COL_L     As Long = ueColL
Public Const COL_M     As Long = ueColM
Public Const COL_N     As Long = ueColN
