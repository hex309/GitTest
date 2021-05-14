Attribute VB_Name = "STEP02"
Option Explicit

Sub STEP02モジュール()

    Call 制御文解析(Pub制御シート名)
    
End Sub

Private Sub 制御文解析Test()
    制御文解析 WSNAME_WARIKOMI
End Sub

Public Sub 制御文解析(ByVal 制御シート名 As String)

    '再帰処理付プロシージャ（ご注意を！）

    Dim MaxRow As Long
    
    MaxRow = ThisWorkbook.Sheets(制御シート名).Cells(Rows.Count, 10).End(xlUp).Row
    
    Dim ウィンドウ As String, アクション As String, タグ As String, 型式 As String, クリック対象 As String, オプション1 As String, オプション2 As String
    Dim テキスト As String

    Dim i As Long
    Dim 対象ファイルパス As String
'    Dim 対象ファイル名 As String
    Dim CSVFile As Workbook
    For i = 20 To MaxRow
        
        DoEvents
'        If fStop = True Then GoTo HandleError
        
        '----------------------------------------------------
        'レコード選択判定
        '----------------------------------------------------
        If ThisWorkbook.Sheets(制御シート名).Cells(i, 7).Value <> "○" Then
            If ThisWorkbook.Sheets(制御シート名).Cells(i, 8).Value <> "○" Then GoTo continue
        End If
        
        '----------------------------------------------------
        'データ取得
        '----------------------------------------------------
        Pub操作番号 = ThisWorkbook.Sheets(制御シート名).Cells(i, 9).Value
        
        ウィンドウ = ThisWorkbook.Sheets(制御シート名).Cells(i, 10).Value
        アクション = ThisWorkbook.Sheets(制御シート名).Cells(i, 11).Value
        タグ = ThisWorkbook.Sheets(制御シート名).Cells(i, 12).Value
        型式 = ThisWorkbook.Sheets(制御シート名).Cells(i, 13).Value
        クリック対象 = ThisWorkbook.Sheets(制御シート名).Cells(i, 14).Value
        オプション1 = ThisWorkbook.Sheets(制御シート名).Cells(i, 15).Value
        オプション2 = ThisWorkbook.Sheets(制御シート名).Cells(i, 16).Value
        
        テキスト = ThisWorkbook.Sheets(制御シート名).Cells(i, 20).Value
        
        '----------------------------------------------------
        'アクション判定（停止）
        '----------------------------------------------------
        If アクション = "停止" Then MessageBox 0, "停止アクション", "オートパイロット停止", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION: End
        
        '----------------------------------------------------
        'アクション判定（データ取得）
        '----------------------------------------------------
        If アクション = "データ取得" Then
            Pub対象ファイルパス = GetElementByID(ウィンドウ, クリック対象)
'            Pub対象ファイルパス = "C:\Users\21501173\Desktop\NESIC様\PSI熱海 ネットワークインフラ構築_【Wi-Fi構築・保守運用】.csv"
            If Pub対象ファイルパス <> "" Then
                Set CSVFile = Workbooks.Open(Pub対象ファイルパス)
                Pub見積登録件名 = CSVFile.Worksheets(1).Range("K3").Value
                Pub件名 = Pub見積登録件名
                Pub営業者コード = CSVFile.Worksheets(1).Range("D3").Value
'                Pub主任者コード = CSVFile.Worksheets(1).Range("E3").Value
                Pub主任者コード = CSVFile.Worksheets(1).Range("P3").Value
                Pub工期FROM = CSVFile.Worksheets(1).Range("M3").Value
                Pub工期TO = CSVFile.Worksheets(1).Range("N3").Value
                Pub見積前提条件 = TextFileData(GetTextFileFullPath(Pub対象ファイルパス))
                
                SetKengyoHo CSVFile.Worksheets(1)
                SetChotatsukubun CSVFile.Worksheets(1)
                SetNextPage
                
                CSVFile.Close False
                GoTo continue
            End If
        End If
        '----------------------------------------------------
        'アクション判定（変数入力）
        '----------------------------------------------------
        If アクション = "変数入力" Then
            
            テキスト = IIf(テキスト = "▼件名", Pub見積登録件名, テキスト)
            テキスト = IIf(テキスト = "▼営業者コード", Pub営業者コード, テキスト)
            If テキスト = "▼主任者コード" Then
                If IsEmpty(Pub工期FROM) = True Or IsNull(Pub工期FROM) = True Or Pub工期FROM = "" Then
                    テキスト = vbNullString
                Else
                    テキスト = IIf(テキスト = "▼主任者コード", Pub主任者コード, テキスト)
                End If
            End If
            テキスト = IIf(テキスト = "▼工期FROM", Format(Pub工期FROM, "ddddd"), テキスト)
            テキスト = IIf(テキスト = "▼工期TO", Format(Pub工期TO, "ddddd"), テキスト)
            テキスト = IIf(テキスト = "▼見積前提条件", Pub見積前提条件, テキスト)
            If テキスト = "" Then GoTo continue

        End If
        
        '----------------------------------------------------
        'アクション判定
        '----------------------------------------------------
        If アクション = "新規IE" Then
            Call 新規IE開く(ウィンドウ)
            IE前面 ウィンドウ
            GoTo continue
        End If
        If アクション = "既存IE" Then
            Call 既存IE開く(ウィンドウ)
            IE前面 ウィンドウ
            GoTo continue
        End If
        If アクション = "LUCASログイン情報入力" Then LucasLoginForm.Show: GoTo continue
        If アクション = "LUCASログイン" Then
            Call ieInTextBoxTagInputTypeText(ウィンドウ, アクション, "input", "USER_INPUT", オプション1, オプション2, PubClsLucasAuth.LucasID)
            Call ieInTextBoxTagInputTypeText(ウィンドウ, アクション, "input", "PASSWORD", オプション1, オプション2, PubClsLucasAuth.LucasPW)
            Call ieClickSubmitButtonTagInputTypeSubmit(ウィンドウ, アクション, "input", "ログイン", オプション1, オプション2): GoTo continue
        End If
        
        
        
        If アクション = "ログイン確認" Then Call ieCheckSSISLogin(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
        
        If アクション = "見積出力" Then Call 例外_見積出力.例外_見積出力モジュール: GoTo continue
        
        If アクション = "通知バー" Then Call ieDownloadFileNbOrDlg(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
        
        '        If アクション = "値抽出" Then
        '            If 型式 = "hidden" Then Call ieExValueTagInputTypeHidden(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
        '        End If
        
        '----------------------------------------------------
        'アクション判定（割込）
        '----------------------------------------------------
        If アクション = "割込" Then
            Call 制御文解析(クリック対象) '再帰処理
            IE前面 ウィンドウ
        End If
         
        '----------------------------------------------------
        'タグ判定／形式判定
        '----------------------------------------------------
        If タグ = "a" Then
            If i = MaxRow Then
                Pub見積番号 = 見積番号取得
            End If
            Call ieClickLinkTagAhrefTypeNone(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2): GoTo continue
            IE前面 ウィンドウ
        End If
        
        If タグ = "input" Then
            If 型式 = "button" Then Call ieClickButtonTagInputTypeButton(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2): GoTo continue
            If 型式 = "checkbox" Then Call ieClickCheckBoxTagInputTypeCheckBox(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2): GoTo continue
            If 型式 = "text" Then Call ieInTextBoxTagInputTypeText(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
            If 型式 = "file" Then Call ieClickButtonTagInputTypeFile(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
            If 型式 = "submit" Then Call ieClickSubmitButtonTagInputTypeSubmit(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2): GoTo continue
            If 型式 = "radio" Then Call ieClickRadioButtonTagInputTypeRadio(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2): GoTo continue
            If 型式 = "password" Then Call ieInPasswordBoxTagInputTypePassword(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
            IE前面 ウィンドウ
        End If
        
        If タグ = "textarea" Then
            If 型式 = "text" Then Call ieInTextTagTextAreaTypeText(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
            IE前面 ウィンドウ
        End If
        

        If タグ = "select" Then
            Call ieClickSelectBoxTagSelect(ウィンドウ, アクション, タグ, クリック対象, オプション1, オプション2, テキスト): GoTo continue
            IE前面 ウィンドウ
        End If
continue:

    Next
    
    Exit Sub
    
HandleError:

    MessageBox 0, "オートパイロットは停止しました。", "確認", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION
    
    End
End Sub

