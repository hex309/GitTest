Attribute VB_Name = "HTMLDocumentHelper"
Option Explicit

'https://www.ka-net.org/blog/?p=13587

Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal pString As LongPtr, ByRef pCLSID As Currency) As Long
Private Declare PtrSafe Function RegisterWindowMessageW Lib "user32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function SendMessageTimeoutW Lib "user32" (ByVal hWnd As LongPtr, ByVal msg As Long, ByVal wParam As LongPtr, ByRef lParam As LongPtr, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As LongPtr
Private Declare PtrSafe Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, ByRef riid As Currency, ByVal wParam As LongPtr, ppvObject As Any) As Long
Private Enum SMTO
    NORMAL = 0
    BLOCK = 1
    ABORTIFHUNG = 2
    NOTIMEOUTIFNOTHUNG = 8
End Enum

' Internet Explorer_Server ウィンドウのハンドルから HTMLDocument オブジェクトを取得する
'
' 第 1 引数: InternetExplorer_Server のウィンドウハンドル
' 第 2 引数: 省略可能(タイムアウト時間)
' 第 3 引数: 省略可能(1:IHTMLDocument～8:IHTMLDocument8)
Public Function GetHtmlDocument(ByVal hWnd_InternetExplorer_Server As LongPtr _
    , Optional ByVal uTimeout As Long = 1000 _
    , Optional ByVal documentVersion As Integer = 1) As Object  ' As MSHTML.IHTMLDocument
    Set GetHtmlDocument = Nothing
    
    If documentVersion <= 0 Then
        documentVersion = 1
    ElseIf documentVersion >= 8 Then
        documentVersion = 8
    End If
    Dim IID_IHTMLDocumentX As String
    IID_IHTMLDocumentX = Split(",{626FC520-A41E-11cf-A731-00A0C9082637},{332c4425-26cb-11d0-b483-00c04fd90119},{3050f485-98b5-11cf-bb82-00aa00bdce0b},{3050f69a-98b5-11cf-bb82-00aa00bdce0b},{3050f80c-98b5-11cf-bb82-00aa00bdce0b},{30510417-98b5-11cf-bb82-00aa00bdce0b},{305104b8-98b5-11cf-bb82-00aa00bdce0b},{305107d0-98b5-11cf-bb82-00aa00bdce0b}", ",")(documentVersion - 1)
    Dim InterfaceId(1) As Currency
    Call CLSIDFromString(StrPtr(IID_IHTMLDocumentX), InterfaceId(0))
    
    Dim lngMsg As Long
    lngMsg = RegisterWindowMessageW(StrPtr("WM_HTML_GETOBJECT"))
    If lngMsg <> 0 Then
        Dim lpdwResult As Long
        If SendMessageTimeoutW(hWnd_InternetExplorer_Server, lngMsg, 0, 0, SMTO.ABORTIFHUNG, uTimeout, lpdwResult) <> 0 Then
            Dim hResult As Long
            hResult = ObjectFromLresult(lpdwResult, InterfaceId(0), 0, GetHtmlDocument)
            If hResult <> 0 Then
                Err.Raise hResult
            End If
        End If
    End If
End Function
