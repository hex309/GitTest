Attribute VB_Name = "Main"
Option Explicit

#Const cnsTest = 0   '#本番
'#Const cnsTest = 1     '#テスト
'本番/テストを切り替える場合は、アカウントシートの定数も書き換えること！

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const LOCK_PSWD_RNG As String = "B3"

Private Const CAP_SUB_RNG As String = "E4"
Private Const FIN_SUB_RNG As String = "E11"

Public Const AC_CORP_IDX As Long = 1
Public Const AC_SITE_IDX As Long = 2
Public Const AC_ADDR_IDX As Long = 3
Public Const AC_COID_IDX As Long = 4
Public Const AC_ACNT_IDX As Long = 5
Public Const AC_PSWD_IDX As Long = 6
Public Const AC_DLLO_IDX As Long = 7
Public Const AC_EDLO_IDX As Long = 8
Public Const AC_IWUL_IDX As Long = 9
Public Const AC_IWDL_IDX As Long = 10
Public Const AC_IWCO_IDX As Long = 11

Public Const SC_CORP_IDX As Long = 2
Public Const SC_MLFL_IDX As Long = 8
Public Const SC_LAST_UPDT_COL_IDX As Long = 11

Public Const NO_USER_MSG As String = "(候補者なし)"

Public Const OUT_DATA_ERR As Long = 1 + vbObjectError + 512
Public Const GET_CUR_REG_ERR As Long = 2 + vbObjectError + 512

Public opeLog As Collection
Public mailInfo As Collection
Public cancelFlg As Boolean

Enum dataType
    personal = 1
    Seminar = 2
End Enum

Sub Main()
    Dim scRng As Range '「メイン」シート表開始セル(「項番」)
    Dim initTime As Date
    Dim finMailMsg As String
    Dim i As Long
    
    '3つの前処理
    '①「実行者氏名」のデータチェック、②対象シートの保護チェック、③対象シートのフィルター解除
    If Not preCheck Then
        Exit Sub
    End If
    
    cancelFlg = False
    initTime = Now()
      
    Set opeLog = New Collection  'コレクション使用

    '「メインシート」の表データを取得
    Set scRng = getCurrentRegion(ScenarioSh.Cells(2, 1), 1, False)
    
    On Error Resume Next
    '「実行ログ」シートの表データを取得し、セルの値をクリア
    getCurrentRegion(LogSh.Cells(2, 1), 1, False).ClearContents
    On Error GoTo 0
    
'###### 20200602 AIM：Yamamoto ######
'対象企業名に使用禁止文字がないか判定
    If corpNameCheck(scRng) Then
#If cnsTest = 1 Then
        MsgBox "対象企業名に使用禁止文字なし"
        Exit Sub
#End If
    Else
        MsgBox "対象企業名に使用禁止文字が含まれています。"
        Exit Sub
    End If
'####################################
    
    For i = 1 To scRng.Rows.Count  '10行目まで(項番が10迄入力されている)
        '完了メールを初期化
        Set mailInfo = New Collection  'コレクション使用
        
        ' 「メイン」シートの項目名「実行」がTRUEか判定
        '【要変更】列番号は定数に変更したほうが良さそう
        If scRng(i, 7).Value Then
            'TRUEの場合　（下記どちらも、「Worksheet_Change」が実行される）
            scRng(i, 9).Value = initTime  '項目名「開始日時」に上記で取得したNowを入力
            scRng(i, 10).Value = vbNullString  '項目名「処理結果」を空欄にする

            If excProcess(scRng(i, 2).Value) Then
                scRng(i, 10).Value = "OK"
                
                If ScenarioSh.getMyNaviOmit Or ScenarioSh.getRikuNaviOmit Or ScenarioSh.getPsOmit Or ScenarioSh.getSmOmit Then
                    opeLog.Add "処理の無効がTRUEのため、終了日時は更新しません。" & vbCrLf & _
                               "今回の更新日時：" & initTime
                    scRng(i, 11).offset(0, 1).Value = "*"
                Else
                    scRng(i, 11).Value = initTime
                End If
                
                '更新した差分ファイルを更新
                SettingSh.ensureOldPath scRng(i, 2).Value
                
                finMailMsg = scRng(i, 2).Value & "のi-Webインポートが完了しました！" & vbCrLf & vbCrLf & getMailInfo()
                
                If sendFinAlert(finMailMsg, scRng(i, 2).Value) Then
                    opeLog.Add scRng(i, 2).Value & "のi-Webインポートが完了しました！"
                    outputLog "Import Completed.", True, scRng(i, 2).Value, vbNullString
                Else
                    opeLog.Add "完了のメール通知が送られておりません。"
                    outputLog "Finish Alert Mail could not send.", True, scRng(i, 2).Value, vbNullString
                End If
            Else
                scRng(i, 10).Value = "NG"
            End If
            
        End If
    Next
    
    If Not LogSh.AutoFilterMode Then LogSh.setAutoFilter
    If Not OldLogSh.AutoFilterMode Then OldLogSh.setAutoFilter

    Unload AlertBox
    Set opeLog = Nothing
    
End Sub

'3つの前処理
'①実行者氏名のデータ有無チェック、②シート保護チェック、③フィルターチェック
Public Function preCheck() As Boolean
    Dim sh As Variant
    Dim name As String
    
    name = ScenarioSh.getUserName
    
    '「メイン」シートの「実行者氏名」データ有無チェック
    If name = vbNullString Or name = NO_USER_MSG Then
        MsgBox "実行者名が空白です。" & vbCrLf & "実行者名を選択したのち、再度実行してください。"
        Exit Function
    End If

    '4つのシートのシート保護チェック(①「メイン」、②「アカウント」、③「過去ログ」、④「メールアカウント」シート)
    'シート保護が掛けられていない場合はアラート表示
    For Each sh In Array(ScenarioSh, AccountSh, OldLogSh, MailSettingSh)
        If Not sh.ProtectContents Then
            MsgBox sh.name & "シートが保護解除中です。" & vbCrLf & "保護を再開したのち、再度実行してください。"
            Exit Function
        End If
    Next
    
    '2つのシートにフィルターが掛かっている場合はフィルター解除(①「実行ログ」、②「過去ログ」シート)
    If LogSh.AutoFilterMode Then LogSh.setAutoFilter
    If OldLogSh.AutoFilterMode Then OldLogSh.setAutoFilter
    
    preCheck = True

End Function

'##### 20200602 AIM：Yamamoto #####
'対象企業名に使用禁止文字が含まれていた場合、ログに追記
Private Function corpNameCheck(scRng As Range) As Boolean
    Dim i As Long
    Dim re As New RegExp
    
    corpNameCheck = True

    For i = 1 To scRng.Rows.Count
        With re
            .Global = True
            .Pattern = "([!#$%&'""`+\-/=~,;:@^<>\?\*\|\{\}\(\)\[\]\\])+"
            If .test(scRng.Cells(i, 2).Value) Then
                corpNameCheck = False
                opeLog.Add "対象企業名に使用禁止文字が含まれています。"
                outputLog "Corporate Name Check ", False, scRng.Cells(i, 2).Value, vbNullString
            End If
        End With
    Next
End Function

Private Function excProcess(ByVal tgtCorpName As String) As Boolean

    AlertBox.Caption = tgtCorpName & "実行中"
    AlertBox.Show False

'###IEラッパーを起動
    '##i-Web用のIEラッパーを起動
    Dim iweb As CorpSite
    Set iweb = New CorpSite
    
    On Error GoTo wrapperErr
    'ログ追加
    opeLog.Add "InternetExploreを起動中..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr
    iweb.cleanUpTgtSite

    '##マイナビ用のIEラッパーを起動
    Dim myNavi As CorpSite
    Dim myNaviFlg As Boolean
    Set myNavi = New CorpSite

    If myNavi Is Nothing Then GoTo wrapperErr
    If myNavi.setCorp(tgtCorpName, "マイナビ") Then
        myNaviFlg = True
        myNavi.cleanUpTgtSite
    End If
      
    '##リクナビ用のIEラッパーを起動
    Dim rikuNavi As CorpSite
    Dim rikuNaviFlg As Boolean
    Set rikuNavi = New CorpSite

    If rikuNavi Is Nothing Then GoTo wrapperErr
    If rikuNavi.setCorp(tgtCorpName, "リクナビ") Then
        rikuNaviFlg = True
        rikuNavi.cleanUpTgtSite
    End If
    
    '##メインシートのマイナビ無効、リクナビ無効を反映させる
    If ScenarioSh.getMyNaviOmit Then myNaviFlg = False
    If ScenarioSh.getRikuNaviOmit Then rikuNaviFlg = False
    
    If Not myNaviFlg And Not rikuNaviFlg Then
        opeLog.Add "実行可能なアカウントがないため処理を終了します。"
        GoTo wrapperErr
    End If
    
    Dim psFlg As Boolean: psFlg = True
    Dim smFlg As Boolean: smFlg = True
    
    '##メインシートのマイナビ無効、リクナビ無効を反映させる
    If ScenarioSh.getPsOmit Then psFlg = False
    If ScenarioSh.getSmOmit Then smFlg = False
    
    If Not psFlg And Not smFlg Then
        opeLog.Add "実行可能な処理がないため処理を終了します。"
        GoTo wrapperErr
    End If
        
    'ログ出力
    opeLog.Add "起動完了。"
    outputLog "InternetExplore Wrapper startup", True, tgtCorpName, vbNullString
    On Error GoTo 0

'####i-webから個人情報を全件ダウンロード
    Dim iWebPsDataFilePath As String
    
    If psFlg Then
        opeLog.Add "i-webから個人情報を全件ダウンロード開始"
    
        On Error GoTo dlErr
#If cnsTest = 0 Then
        iWebPsDataFilePath = DLiWebData(iweb, iweb.iWebPDLayNo)

        iWebPsDataFilePath = moveFileAddHeadder(iWebPsDataFilePath, "【" & tgtCorpName & "】" & "i-Web個人情報全件_個人情報UL前_")

        If iWebPsDataFilePath = vbNullString Then GoTo dlErr
#End If
        opeLog.Add "全件ダウンロード完了"
    
        On Error GoTo 0
        
    Else
        opeLog.Add "個人情報インポート無効のため、i-webから個人情報を全件ダウンロードをスキップします。"
    
    End If
    
    outputLog "Download i-Web personal data", True, tgtCorpName, iWebPsDataFilePath
    

'###マイナビから個人情報/セミナー情報をダウンロード
    Dim myNavPsDataFilePath As String
    Dim myNavPsDataFileName As String
    Dim myNavPsFlg As Boolean
    Dim mynavSmDataFilePath As String
    Dim myNavSmDataFileName As String
    Dim myNavSmFlg As Boolean

    '2018/07 以降テストでのマイナビログインは禁止！必要であればお客様の許可を得る事！

    If Not myNaviFlg Then
        opeLog.Add "マイナビの処理はありません。"
        outputLog "Skip Mynavi Download/Upload", True, tgtCorpName, vbNullString
        GoTo MY_NAV_SKIP
    End If

#If cnsTest = 1 Then
    '#テスト用ダミーデータ
    If Not psFlg Then
        opeLog.Add "個人情報インポート無効のため、マイナビからの個人情報ダウンロードをスキップします。"
        outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
    Else
        myNavPsFlg = True
        myNavPsDataFilePath = getDlFilePath("20000シンカ_マイ_個人.csv")
    End If
    
    If Not smFlg Then
        opeLog.Add "セミナーインポート無効のため、マイナビからのセミナー情報ダウンロードをスキップします。"
        outputLog "Skip MyNavi Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        myNavSmFlg = True
        'mynavSmDataFilePath = getDlFilePath("20000シンカ_セミナ_個人.csv")
        mynavSmDataFilePath = "C:\Users\11402086\Desktop\test\2\20000シンカ_マイ_セミナ2.csv"
        SettingSh.Cells(11, 2).Value = "C:\Users\11402086\Desktop\test\2\20000シンカ_マイ_セミナ.csv"
    End If
       
    GoTo MY_NAV_SKIP
#End If

    '##マイナビログイン
    AlertBox.Label1 = "マイナビにログイン中.."
    opeLog.Add "マイナビにログイン中.. アカウント：" & myNavi.userName

    On Error GoTo loginErr
    If Not loginMyNavi(myNavi) Then GoTo loginErr
    On Error GoTo 0

    opeLog.Add "マイナビにログイン完了"
        
    '##マイナビ個人情報DL予約
    If Not psFlg Then
        opeLog.Add "個人情報インポート無効のため、マイナビからの個人情報ダウンロードをスキップします。"
        outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
    Else
        '##更新日時を設定して検索
        AlertBox.Label1 = "更新日時を設定して検索中.."

        opeLog.Add "マイナビの個人情報検索ページに移動中.."
        On Error GoTo pageErr
        If Not moveSearchWindow(myNavi) Then GoTo pageErr
        On Error GoTo 0

        opeLog.Add "移動完了。個人情報検索開始"
        
        myNavPsFlg = myNavi.dlLayout <> vbNullString
        
        If Not myNavPsFlg Then
            opeLog.Add "ナビサイトレイアウト名（個人情報Download用）がアカウントシートに記載されていません。スキップします。"
            outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchMyNaviDt(myNavi, dataType.personal) Then
                opeLog.Add "新規データなし。"
                outputLog "Did not hit new MyNavi Personal Data", True, tgtCorpName, vbNullString
                
                myNavPsFlg = False
            Else
                '#CSVファイルDLを予約
                AlertBox.Label1 = "マイナビで個人情報CSVを作成中.."
                opeLog.Add "マイナビで個人情報CSVを作成中.."
    
                myNavPsDataFileName = makeMyNaviDt(myNavi, dataType.personal)
                If myNavPsDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
    End If
        
    '##マイナビセミナー情報DL予約
        
    myNavSmFlg = False
        
    If Not smFlg Then
        opeLog.Add "セミナーインポート無効のため、マイナビからのセミナー情報ダウンロードをスキップします。"
        outputLog "Skip MyNavi Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# イベント情報を全検索
        AlertBox.Label1 = "セミナー情報を全検索中.."
        opeLog.Add "マイナビでセミナー情報の検索開始.."
        
        On Error GoTo dlErr

        myNavSmFlg = myNavi.dlLayoutEV <> vbNullString

        If Not myNavSmFlg Then
            opeLog.Add "ナビサイトレイアウト名（イベント・セミナー情報Download用）がアカウントシートに記載されていません。スキップします。"
            outputLog "Skip MyNavi Seminar Data Download ", True, tgtCorpName, vbNullString
        Else
            opeLog.Add "マイナビのセミナー情報検索ページに移動中.."

            On Error GoTo pageErr
            If Not moveSearchWindow(myNavi, Not psFlg) Then GoTo pageErr
            On Error GoTo 0

            If Not searchMyNaviDt(myNavi, dataType.Seminar) Then
                If InStr(opeLog(opeLog.Count), "データはありませんでした。") = 0 Then
                    GoTo dlErr
                Else
                    opeLog.Add "新規データなし。"
                    outputLog "Did not hit new MyNavi Seminar Data", True, tgtCorpName, vbNullString
                    myNavSmFlg = False
                End If
            Else
                '#CSVファイルDLを予約
                AlertBox.Label1 = "マイナビでセミナー情報CSVを作成中.."
                opeLog.Add "マイナビでセミナー情報CSVを作成中.."

                myNavSmDataFileName = makeMyNaviDt(myNavi, dataType.Seminar)
                If myNavSmDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
        
        On Error GoTo 0
    End If
    
    '# CSVファイルをダウンロード

    If myNavPsFlg Then chkDateCreated myNavi, myNavPsDataFileName
    If myNavSmFlg Then chkDateCreated myNavi, myNavSmDataFileName
    
    If myNavPsFlg Then
        opeLog.Add "個人情報データダウンロード開始.."
        myNavPsDataFilePath = dlMyNaviCSV(myNavi, myNavPsDataFileName)
        If myNavPsDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "個人情報データダウンロード完了。"
        outputLog "Download myNavi personal CSV data", True, tgtCorpName, myNavPsDataFilePath
    End If
    
    If myNavSmFlg Then
        opeLog.Add "セミナーデータダウンロード開始。"
        mynavSmDataFilePath = dlMyNaviCSV(myNavi, myNavSmDataFileName)
        If mynavSmDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "セミナーデータダウンロード完了。"
        outputLog "Download myNavi Seminar CSV data", True, tgtCorpName, mynavSmDataFilePath
    End If

MY_NAV_SKIP:

' ★追加分★★★★★★★★★★★★★★★
    If myNavSmFlg Then
        mynavSmDataFilePath = getDiffFile(mynavSmDataFilePath, tgtCorpName, myNavi.lastUpdate)
        If mynavSmDataFilePath = vbNullString Then GoTo diffErr
    End If
' ★★★★★★★★★★★★★★★★★★★

'###マイナビサイトの表示終了
    On Error Resume Next
    myNavi.visible False
    On Error GoTo 0

'###リクナビから個人情報/セミナー情報をダウンロード
    Dim rkNavPsDataFilePath As String
    Dim rkNavPsDataFileName As String
    Dim rkNavPsFlg As Boolean
    Dim rknavSmDataFilePath As String
    Dim rkNavSmDataFileName As String
    Dim rkNavSmFlg As Boolean

    If Not rikuNaviFlg Then
        opeLog.Add "リクナビの処理はありません。"
        outputLog "Skip RikuNavi Download/Upload", True, tgtCorpName, vbNullString
        GoTo RIKU_NAV_SKIP
    End If

#If cnsTest = 1 Then
    If Not psFlg Then
        opeLog.Add "個人情報インポート無効のため、リクナビからの個人情報ダウンロードをスキップします。"
        outputLog "Skip RikuNaviSeminar CSV data ", True, tgtCorpName, vbNullString
    Else
        rkNavPsFlg = True
        rkNavPsDataFilePath = getDlFilePath("20000シンカ_リク_個人.csv")
    End If
    
    If Not smFlg Then
        opeLog.Add "セミナーインポート無効のため、リクナビからのセミナー情報ダウンロードをスキップします。"
        outputLog "Skip RikuNavi  Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        rkNavSmFlg = True
        'rknavSmDataFilePath = getDlFilePath("20000シンカ_リク_セミナ.csv")
        rknavSmDataFilePath = "C:\Users\11402086\Desktop\test\2\20000シンカ_リク_セミナ.csv"
    End If
        
    GoTo RIKU_NAV_SKIP
#End If
    
    
    '# リクナビログイン
    AlertBox.Label1 = "リクナビにログイン中.."
    opeLog.Add "リクナビにログイン中.. アカウント：" & rikuNavi.userName

    On Error GoTo loginErr
    If Not loginRikuNavi(rikuNavi) Then GoTo loginErr
    On Error GoTo 0

    opeLog.Add "リクナビにログイン完了"

    '##リクナビ個人情報DL予約
    If Not psFlg Then
        opeLog.Add "個人インポート無効のため、リクナビからの個人情報ダウンロードをスキップします。"
        outputLog "Skip RikuNaviSeminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# 更新日時を設定して新規登録された個人情報検索
        AlertBox.Label1 = "更新日時を設定して検索中.."
        opeLog.Add "リクナビで個人情報の検索開始.."

        On Error GoTo dlErr

        rkNavPsFlg = rikuNavi.dlLayout <> vbNullString

        If Not rkNavPsFlg Then
            opeLog.Add "ナビサイトレイアウト名（個人情報Download用）がアカウントシートに記載されていません。スキップします。"
            outputLog "Skip RikuNavi Personal Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchRikuNaviDt(rikuNavi, dataType.personal) Then
                If InStr(opeLog(opeLog.Count), "で検索しましたが該当するデータはありませんでした。") = 0 Then
                    GoTo dlErr
                Else
                    opeLog.Add "個人情報の新規登録なし。"
                    outputLog "RikuNavi : Did not hit new Personal Data", True, tgtCorpName, vbNullString
    
                    rkNavPsFlg = False
                End If
            Else
                rkNavPsDataFileName = makeRikuNaviDt(rikuNavi, dataType.personal)
                If rkNavPsDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
        
        On Error GoTo 0
    End If

    '##リクナビセミナー情報DL予約
    rkNavSmFlg = False
    
    If Not smFlg Then
        opeLog.Add "セミナーインポート無効のため、リクナビからのセミナー情報ダウンロードをスキップします。"
        outputLog "Skip RikuNavi  Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# イベント情報を全検索
        AlertBox.Label1 = "セミナー情報を全検索中.."
        opeLog.Add "リクナビでセミナー情報の検索開始.."

        rkNavSmFlg = rikuNavi.dlLayoutEV <> vbNullString

        If Not rkNavSmFlg Then
            opeLog.Add "ナビサイトレイアウト名（イベント・セミナー情報Download用）がアカウントシートに記載されていません。スキップします。"
            outputLog "Skip RikuNavi Seminar Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchRikuNaviDt(rikuNavi, dataType.Seminar) Then
                opeLog.Add "セミナー情報の登録なし。"
                outputLog "RikuNavi : Did not hit new Seminar Data", True, tgtCorpName, vbNullString

                rkNavSmFlg = False
            Else
                rkNavSmDataFileName = makeRikuNaviDt(rikuNavi, dataType.Seminar)
                If rkNavSmDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
    End If

    '# CSVの作成を待機
    If rkNavPsFlg Then
        AlertBox.Label1 = "リクナビで個人情報用のCSVを作成中（時間がかかります）.."
        opeLog.Add "リクナビで個人情報用のCSVを作成中.."
        waitRikuNaviCSV rikuNavi, rkNavPsDataFileName
    End If

    If rkNavSmFlg Then
        AlertBox.Label1 = "リクナビでセミナー情報用のCSVを作成中（時間がかかります）.."
        opeLog.Add "リクナビでセミナー情報用のCSVを作成中.."
        waitRikuNaviCSV rikuNavi, rkNavSmDataFileName
    End If

    '# CSVファイルが出揃ったらをDLを開始、DL成功したらパスが返ってくる
    If rkNavPsFlg Then
        opeLog.Add "個人情報ダウンロード開始.."
        rkNavPsDataFilePath = dlRikuNaviCSV(rikuNavi, rkNavPsDataFileName)
        If rkNavPsDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "個人情報ダウンロード完了。"
        outputLog "Download RikuNavi personal data", True, tgtCorpName, rkNavPsDataFilePath
    End If

    If rkNavSmFlg Then
        opeLog.Add "セミナー情報ダウンロード開始.."
        rknavSmDataFilePath = dlRikuNaviCSV(rikuNavi, rkNavSmDataFileName)
        If rknavSmDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "セミナー情報ダウンロード完了。"
        outputLog "Download RikuNavi seminar data", True, tgtCorpName, rknavSmDataFilePath
    End If

'###リクナビサイトの表示終了
    On Error Resume Next
    rikuNavi.visible False
    On Error GoTo 0

RIKU_NAV_SKIP:

    On Error GoTo ulErr
'###マイナビの情報をi-Webへアップロード
    If Not myNavPsDataFilePath = vbNullString And myNavPsFlg And psFlg Then

        AlertBox.Label1 = "マイナビのCSVから、iWebへ個人情報をアップロードしています。"
        opeLog.Add "マイナビのCSVから、iWebへ個人情報をアップロード中.."
        mailInfo.Add "＜個人情報インポート＞" & vbCrLf & "(マイナビ)"

        If Not ULPersonalData(iweb, myNavPsDataFilePath, myNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "アップロード完了"
        outputLog "MyNavi personal data Upload to i-Web", True, tgtCorpName, myNavPsDataFilePath
    Else
        mailInfo.Add "＜個人情報インポート＞" & vbCrLf & "(マイナビ)" & vbCrLf & "対象者がいませんでした。"
    End If

'###リクナビの情報をi-Webへアップロード
    If Not rkNavPsDataFilePath = vbNullString And rkNavPsFlg And psFlg Then

        AlertBox.Label1 = "リクナビのCSVから、iWebへ個人情報をアップロードしています。"
        opeLog.Add "リクナビのCSVから、iWebへ個人情報をアップロード中.."
        mailInfo.Add "(リクナビ)"

        If Not ULPersonalData(iweb, rkNavPsDataFilePath, rikuNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "アップロード完了"
        outputLog "RikuNavi personal data Upload to i-Web", True, tgtCorpName, rkNavPsDataFilePath
    Else
        mailInfo.Add "(リクナビ)" & vbCrLf & "対象者がいませんでした。"
    End If

    On Error GoTo 0
      
    If (myNavSmFlg Or rkNavSmFlg) And smFlg Then
    '####i-webから個人情報を全件ダウンロード(追加分を反映)
        'Dim iWebPsDataFilePath As String
        
        AlertBox.Label1 = "i-webから個人情報を全件ダウンロード開始(追加分を反映).."
        opeLog.Add "i-webから個人情報を全件ダウンロード開始"

        On Error GoTo dlErr
#If cnsTest = 0 Then
        iWebPsDataFilePath = DLiWebData(iweb, iweb.iWebPDLayNo)

        iWebPsDataFilePath = moveFileAddHeadder(iWebPsDataFilePath, "【" & tgtCorpName & "】" & "i-Web個人情報全件_個人情報UL後_")

        If iWebPsDataFilePath = vbNullString Then GoTo dlErr
#End If
        opeLog.Add "全件ダウンロード完了"
        outputLog "Download i-Web personal data", True, tgtCorpName, iWebPsDataFilePath
    
        On Error GoTo 0
        
#If cnsTest = 1 Then
    iWebPsDataFilePath = "C:\Users\11402086\Desktop\test\iwebData.csv"
#End If
       
    '####ファイルからi-Webの個人情報をロード
        Dim iwebPeople As people
    
        Set iwebPeople = getPeople("i-Web", iWebPsDataFilePath)
        
        AlertBox.Label1 = "ファイルからi-Webの個人情報をロード中.."
    
        If iwebPeople Is Nothing Then
            outputLog "Load i-Web personal data from csv file", False, tgtCorpName, iWebPsDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load i-Web personal data from csv file", True, tgtCorpName, iWebPsDataFilePath
        
        End If
    Else
        mailInfo.Add vbCrLf & "＜セミナーインポート＞" & vbCrLf & "対象者がいませんでした。"
    End If

'####ファイルからマイナビのセミナー情報をロード
    Dim myNavSeminor As people
    
     AlertBox.Label1 = "ファイルからマイナビのセミナー情報をロード中.."
    
    If Not mynavSmDataFilePath = vbNullString And myNavSmFlg And smFlg Then
        Set myNavSeminor = getPeople("マイナビ", mynavSmDataFilePath, False, iweb.lastUpdate)
    
        If myNavSeminor Is Nothing Then
            outputLog "Load MyNavi seminar data from csv file", False, tgtCorpName, mynavSmDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load MyNavi seminar data from csv file", True, tgtCorpName, mynavSmDataFilePath
        End If
    End If


'####ファイルからリクナビのセミナー情報をロード
    Dim rkNavSeminar As people
    
    AlertBox.Label1 = "ファイルからリクナビのセミナー情報をロード中.."
    
    If Not rknavSmDataFilePath = vbNullString And rkNavSmFlg And smFlg Then
        Set rkNavSeminar = getPeople("リクナビ", rknavSmDataFilePath, False, iweb.lastUpdate)
    
        If rkNavSeminar Is Nothing Then
            outputLog "Load RikuNavi seminar data from csv file", False, tgtCorpName, rknavSmDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load RikuNavi seminar data from csv file", True, tgtCorpName, rknavSmDataFilePath
        End If
    End If
    
'###セミナー情報アップロード
    
    If (myNavSmFlg Or rkNavSmFlg) And smFlg Then
        AlertBox.Label1 = "セミナー情報をアップロード中.."
    
        If Not ULAllSeminarData(iweb, myNavi, myNavSeminor, rkNavSeminar, iwebPeople) Then GoTo ulErr
        outputLog "Uupload all Seminar data", True, tgtCorpName, vbNullString
    Else
        'Do nothing
    End If
    
'###処理完了
    excProcess = True

normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not rikuNavi Is Nothing Then rikuNavi.quitAll
    If Not myNavi Is Nothing Then myNavi.quitAll
    
Exit Function

'###エラー処理
wrapperErr:
    outputLog "IE Wrapper could not startup", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin

loginErr:
    outputLog "Login failed", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin

pageErr:
    outputLog "Failed to navigate the target page", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
csvErr:
    outputLog "CSV could not be created on Navi site", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
dlErr:
    outputLog "Failed to download the CSV file", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
diffErr:
    outputLog "Failed to extract difference", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
ulErr:
    outputLog "Failed to upload the data", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
End Function

Private Function outputLog(ByVal opeName As String, ByVal successFlg As Boolean, ByVal tgtCorpName As String, Optional ByVal tgtFilePath As String)
    Dim tgtCell As Range
    Dim result As String
    Dim errLog As Variant
    Dim errLogs As Collection
    Dim i As Long
    Dim j As Long
    Dim maxLine As Long
    Dim userName As String
    Dim nextFlg As Boolean
    
    Set errLogs = New Collection
    
    If successFlg Then
        result = "成功"
    Else
        result = "失敗"
    End If
    
    maxLine = SettingSh.getLogMaxLow
    
    If maxLine = 0 Or maxLine > 253 Then maxLine = 253
    
    j = 1
    
    '1ログで32,767文字を超えるものはない（意図的に作らないとできない）
    For i = 1 To opeLog.Count
        'ログをまとめる
        errLog = errLog & IIf(j > 1, vbCrLf, vbNullString) & opeLog(i)
        j = j + 1
        
        'ログが指定行を超える、もしくは最終ログのときフラグを立てる。
        If j = maxLine + 1 Or i = opeLog.Count Then
            nextFlg = True
        
        'ログが最終ログでなく、次のログと合わせて32,767文字を超えるときフラグを立てる。
        ElseIf Len(errLog) + Len(opeLog(i + 1)) + 2 > 32767 Then
            nextFlg = True
        End If
        
        'フラグが立っていたら、ログ'sに追加して、いったんクリア。
        If nextFlg Then
            errLogs.Add errLog
            errLog = vbNullString
            j = 1
            nextFlg = False
        End If
    Next
    
    'opeLog クリア
    Set opeLog = New Collection
    
    userName = ScenarioSh.getUserName
    
    For Each errLog In errLogs
        With LogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0)
            .Value = Now()
            .offset(0, 1).Value = result
            .offset(0, 2).Value = tgtCorpName
            .offset(0, 3).Value = opeName
            .offset(0, 4).Value = tgtFilePath
            .offset(0, 5).Value = errLog
            .offset(0, 6).Value = userName
            'OldLogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0).Resize(1, 6).Interior.Color = .Resize(1, 6).Interior.Color
            OldLogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0).Resize(1, 7).Value = .Resize(1, 7).Value
        End With
    Next

End Function

Private Function getMailInfo() As String
    Dim i As Long
    
    For i = 1 To mailInfo.Count
        getMailInfo = getMailInfo & IIf(i > 1, vbCrLf, vbNullString) & mailInfo(i)
    Next
    
    Set mailInfo = New Collection
    
End Function

'「メイン」シート「個人情報データ個別UL」ボタンを押下後、実行
'実行前提として、対象セルが「メイン」シート表の項目名「対象企業名」データを選択している事が条件
'ファイルを指定してもらい、ファイルの対象企業を選択してもらった後、UpLoad処理を実行
Public Sub upPsDataOnly()
    Dim i As Long
    Dim corpName As String '対象企業名
    Dim csvPath As String  '対象CSｖファイルフルパス
    Dim ret As Long
    
    '当モジュール「Main」のプロシージャ「preCheck」は、3つの前処理を実行
    '①実行者氏名のデータ有無チェック、②対象シート(「メイン」「アカウント」「過去ログ」「メールアカウント」シート)シート保護チェック、
    '③対象シート(「実行ログ」「過去ログ」)のフィルター解除
    If Not preCheck Then
        Exit Sub
    End If
    
    cancelFlg = False
      
    Set opeLog = New Collection
    Set mailInfo = New Collection
    
    corpName = Cells(Selection.row, 2)  '対象企業名を取得(対象企業名のセルを選択している前提)
    
    If corpName = vbNullString Then Exit Sub  '対象セルが、「メイン」シートの項目名「対象企業名」データを選択していない場合は、処理終了
    If MsgBox("対象企業は" & corpName & "でよろしいですか？", vbYesNo) <> vbYes Then Exit Sub
    
    MsgBox "個人情報アップロードするCSVファイルを選択してください。"
    'モジュール「FileOpe」の「getFilePathByDialog」の処理を実行
    'CSVファイルを、ユーザーに選択してもらう
    csvPath = getFilePathByDialog("*.csv", "CSVファイル", "個人情報アップロードするファイルを選択してください。")
    If csvPath = vbNullString Then Exit Sub
    
    ret = MsgBox("対象サイトはマイナビですか？" & vbCrLf & "リクナビなら「いいえ(N)」を選択", vbYesNoCancel)
    
    If ret = vbYes Then
        upMyNaviPsDataOnly corpName, csvPath  'マイナビのUpLoad処理へ
    ElseIf ret = vbNo Then
        upRikuNaviPsDataOnly corpName, csvPath  'リクナビのUpLoad処理へ
    Else
        Exit Sub
    End If
    
    opeLog.Add "i-Webインポートが完了しました！"
    
    Dim msg As String
    
    For i = 1 To opeLog.Count
        msg = msg & IIf(msg = vbNullString, vbNullString, vbCrLf) & opeLog(i)
    Next
    
    If Not msg = vbNullString Then
        MsgBox msg, vbInformation
    End If
    
    Set opeLog = Nothing
    
    Unload AlertBox

End Sub

Private Function upRikuNaviPsDataOnly(ByVal tgtCorpName As String, _
                                      ByVal csvPath As String) As Boolean

'###IEラッパーを起動
    '##i-Web用のIEラッパーを起動
    Dim iweb As CorpSite
    Set iweb = New CorpSite
    
    On Error GoTo wrapperErr
    'ログ追加
    opeLog.Add "InternetExploreを起動中..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr
    iweb.cleanUpTgtSite
       
    '##リクナビ用のIEラッパーを起動
    Dim rikuNavi As CorpSite
    Dim rikuNaviFlg As Boolean
    Set rikuNavi = New CorpSite

    If rikuNavi Is Nothing Then GoTo wrapperErr
    If rikuNavi.setCorp(tgtCorpName, "リクナビ") Then
        rikuNaviFlg = True
        rikuNavi.cleanUpTgtSite
    End If
    
   'ログ出力
    opeLog.Add "起動完了。"
    
    On Error GoTo 0
    
    'パス登録
    Dim rkNavPsDataFilePath As String
    Dim rkNavPsFlg As Boolean
    
    opeLog.Add "【手動】リクナビはログインしません。登録されたデータを使います。"
            
    rkNavPsDataFilePath = csvPath
    rkNavPsFlg = True

    On Error GoTo ulErr
'###リクナビの情報をi-Webへアップロード
    If Not rkNavPsDataFilePath = vbNullString And rkNavPsFlg Then

        AlertBox.Label1 = "リクナビのCSVから、iWebへ個人情報をアップロードしています。"
        opeLog.Add "リクナビのCSVから、iWebへ個人情報をアップロード中.."
        mailInfo.Add "(リクナビ)"

        If Not ULPersonalData(iweb, rkNavPsDataFilePath, rikuNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "アップロード完了"
    Else
        opeLog.Add "(リクナビ)" & vbCrLf & "対象者がいませんでした。"
    End If
    On Error GoTo 0
    
normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not rikuNavi Is Nothing Then rikuNavi.quitAll
    
Exit Function

wrapperErr:
    upRikuNaviPsDataOnly = False
    GoTo normalFin
   
ulErr:
    upRikuNaviPsDataOnly = False
    GoTo normalFin

End Function

'引数は2つ
'①対象セルが選択している「メイン」シートの項目名「対象企業名」データ、②選択してもらったCSVファイルフルパス
Private Function upMyNaviPsDataOnly(ByVal tgtCorpName As String, _
                                    ByVal csvPath As String) As Boolean

'###IEラッパーを起動
    '##i-Web用のIEラッパーを起動
    Dim iweb As CorpSite
    Set iweb = New CorpSite  'Indexに「1」を代入して、IEオブジェクトを取得
    
    On Error GoTo wrapperErr
    'ログ追加
    opeLog.Add "InternetExploreを起動中..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr 'iweb取得
    iweb.cleanUpTgtSite
       
    '##マイナビ用のIEラッパーを起動
    Dim myNavi As CorpSite
    Dim myNaviFlg As Boolean
    Set myNavi = New CorpSite

    If myNavi Is Nothing Then GoTo wrapperErr
    If myNavi.setCorp(tgtCorpName, "マイナビ") Then  'マイナビ取得
        myNaviFlg = True
        myNavi.cleanUpTgtSite
    End If
    
   'ログ出力
    opeLog.Add "起動完了。"
    
    On Error GoTo 0
    
    'パス登録
    Dim myNavPsDataFilePath As String
    Dim myNavPsFlg As Boolean
    
    opeLog.Add "【手動】マイナビはログインしません。登録されたデータを使います。"
            
    myNavPsDataFilePath = csvPath
    myNavPsFlg = True

    On Error GoTo ulErr
'###マイナビの情報をi-Webへアップロード
    If Not myNavPsDataFilePath = vbNullString And myNavPsFlg Then

        AlertBox.Label1 = "マイナビのCSVから、iWebへ個人情報をアップロードしています。"
        opeLog.Add "マイナビのCSVから、iWebへ個人情報をアップロード中.."

        If Not ULPersonalData(iweb, myNavPsDataFilePath, myNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "アップロード完了"
    Else
        opeLog.Add "＜個人情報インポート＞" & vbCrLf & "(マイナビ)" & vbCrLf & "対象者がいませんでした。"
    End If
    On Error GoTo 0
    
normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not myNavi Is Nothing Then myNavi.quitAll
    
Exit Function

wrapperErr:
    upMyNaviPsDataOnly = False
    GoTo normalFin
   
ulErr:
    upMyNaviPsDataOnly = False
    GoTo normalFin

End Function

