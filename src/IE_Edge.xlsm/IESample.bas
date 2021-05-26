Attribute VB_Name = "IESample"
Option Explicit

Sub IEOperation()

    Dim ie As IEClass 'クラスの宣言
    Set ie = New IEClass 'クラスの実体化

    Set ie.objIE = New InternetExplorer
    ie.objIE.Visible = True 'True:IEを表示、False:IEを非表示
    ie.objIE.Navigate "https://www.yahoo.co.jp/" 'URLに入れたIEを起動
    Call WaitIE(ie.objIE) 'IEの読み込み待ち関数

    Set ie.htmldoc = ie.objIE.Document '開いたIEのドキュメントをセット

    'IE操作を入れる------------------------


    'IE操作終了---------------------------

    ie.objIE.Quit 'IEを閉じる

End Sub

Function WaitIE(objIE As InternetExplorer) 'IEの読み込み待ち関数

    Do While objIE.Busy = True Or objIE.ReadyState < 4
        DoEvents
    Loop

End Function
