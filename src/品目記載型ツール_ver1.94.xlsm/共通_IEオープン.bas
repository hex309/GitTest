Attribute VB_Name = "共通_IEオープン"
Option Explicit

Public Sub 新規IE開く(ByVal URL As String)
    
    Call ieOpenNew(URL)

End Sub

Public Sub 既存IE開く(ByVal URL As String)

    Call ieOpenExist(URL)

End Sub

Public Sub ieOpenNew(urlName As String, _
           Optional viewFlg As Boolean = True, _
           Optional ieTop As Integer = 30, _
           Optional ieLeft As Integer = 0, _
           Optional ieWidth As Integer = 900, _
           Optional ieHeight As Integer = 900)
    
    'IE(InternetExplorer)のオブジェクトを作成する
    'InternetExplorer.Application は.navigateメソッド実行後に、インスタンスオブジェクトが廃棄される。
    'セキュリティゾーンをまたぐ場合のセッションを維持できないため。
    'Set直後はLowL、保護モードオフの .navigateメソッドは、MidiumLのため、インスタンスを引き継げない
    'Set oPubIE1 = CreateObject("InternetExplorer.Application")　'LowL

    '上記の代案
    Set oPubIE1 = New InternetExplorerMedium
    
    With oPubIE1
        
        'IE(InternetExplorer)を表示・非表示
        .Visible = viewFlg
        
        .Top = ieTop  'Y位置
        .Left = ieLeft  'X位置
        .Width = ieWidth  '幅
        .Height = ieHeight  '高さ
        
        '指定したURLのページを表示する
        .navigate urlName
    
    End With

    'IE(InternetExplorer)が完全表示されるまで待機
    Call ieWaitCheck
    
End Sub

Sub ieOpenExist(ByVal urlName As String)

    Const navOpenInNewTab = &H800

    Dim objShell As Object, objWin As Object
    Dim nFLG As Boolean: nFLG = False

    Set objShell = CreateObject("Shell.Application")

    For Each objWin In objShell.Windows

        '存在しないのに、Internet Explorer が表示されるときがある。このSUBは推奨しない
        If objWin.Name = "Internet Explorer" Then
            Set oPubIE1 = objWin
            nFLG = True
            Exit For
        End If
    Next
    
    If nFLG = True Then
        oPubIE1.Navigate2 urlName, navOpenInNewTab
    Else
        ieOpenNew (urlName)
    End If

End Sub


