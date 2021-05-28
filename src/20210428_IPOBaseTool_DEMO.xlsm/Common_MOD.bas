Attribute VB_Name = "Common_MOD"
Option Explicit
'★共用モジュール
'◆レコードセット取得関連変数
Public Const adOpenKeyset = 1, adLockReadOnly = 1
Public Exl_Cn, Ac_Cn As ADODB.Connection
Public Exl_Rs, Ac_Rs As ADODB.Recordset
Public Ac_Cmd As ADODB.Command
Public str_SQL, str_AcDBcn As String
'◆ブッククローズ用フラグ
Public ClsFlg As Boolean

Public Sub Auto_open()
'★起動時処理
    ClsFlg = False
    Call St_AllUnvis
    ActiveSheet.Unprotect
    ActiveSheet.Range("C2").Select
    Call St_Lock
    Application.WindowState = xlMaximized   'ウィンドウ最大化

End Sub
