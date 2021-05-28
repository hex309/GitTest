Attribute VB_Name = "FileOpe"
Option Explicit
Option Private Module

Public Function csvDivider(ByVal srcCSVpath As String, Optional ByVal maxLineNo As Long = 500) As Collection
    Dim fso As Object 'FileSystemObject
    Dim txtStrm As Object 'TextStream
    Dim titleLine As String
    Dim divText As Object 'TextStream
    Dim newFilePath As String
    Dim i As Long, j As Long
    
    Set csvDivider = New Collection
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(srcCSVpath) Then
        Set txtStrm = fso.OpenTextFile(srcCSVpath)
    Else
        opeLog.Add srcCSVpath & vbCrLf & "上記ファイルがありません。"
        Exit Function
    End If
    
    titleLine = txtStrm.ReadLine
    j = 1
    
    Do Until txtStrm.AtEndOfStream
        If i = 0 Or i >= maxLineNo Then
            Do
                newFilePath = srcCSVpath & "_" & j & ".csv"
                j = j + 1
            Loop While fso.FileExists(newFilePath)
        
            Set divText = fso.CreateTextFile(newFilePath)
            divText.WriteLine titleLine
            i = 1
        End If
    
        divText.WriteLine txtStrm.ReadLine
        i = i + 1
        
        If i = maxLineNo Or txtStrm.AtEndOfStream Then
            csvDivider.Add newFilePath
            
            divText.Close
            Set divText = Nothing
        End If
    Loop

    txtStrm.Close

End Function

Public Function verifyCSV(ByVal srcCSVpath As String) As String
    Dim fso As Object 'FileSystemObject
    Dim txtStrm As Object 'TextStream
    Dim titleLine As String
    Dim dtLine As String
    Dim oldLine As String
    Dim divText As Object 'TextStream
    Dim newFilePath As String
    Dim csvCells As Collection
    Dim modified As Boolean
    Dim i As Long, j As Long
    
    Dim KJ_FAM_NAME_IDX As String
    Dim KJ_FST_NAME_IDX As String
    Dim KN_FAM_NAME_IDX As String
    Dim KN_FST_NAME_IDX As String
    Dim targetIDXs As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(srcCSVpath) Then
        Set txtStrm = fso.OpenTextFile(srcCSVpath)
    Else
        opeLog.Add srcCSVpath & vbCrLf & "上記ファイルがありません。"
        Exit Function
    End If
    
    titleLine = txtStrm.ReadLine
    
    KJ_FAM_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KJ_FAM_NAME")
    KJ_FST_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KJ_FST_NAME")
    KN_FAM_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KN_FAM_NAME")
    KN_FST_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KN_FST_NAME")
    
    targetIDXs = Array(KJ_FAM_NAME_IDX, KJ_FST_NAME_IDX, KN_FAM_NAME_IDX, KN_FST_NAME_IDX)

    Do
        newFilePath = srcCSVpath & "_修正済" & IIf(i = 0, "", "_" & i) & ".csv"
        i = i + 1
    Loop While fso.FileExists(newFilePath)

    Set divText = fso.CreateTextFile(newFilePath)
    divText.WriteLine titleLine
    
    Do Until txtStrm.AtEndOfStream
        dtLine = txtStrm.ReadLine
        oldLine = dtLine
        
        Set csvCells = splitCSVLine(dtLine)
        
        For i = LBound(targetIDXs) To UBound(targetIDXs)
            j = targetIDXs(i)
            
            If j <> 0 Then
                dtLine = Replace(dtLine, csvCells(j), Replace(csvCells(j), " ", vbNullString))
                dtLine = Replace(dtLine, csvCells(j), Replace(csvCells(j), "　", vbNullString))
            End If
        Next
        
        If oldLine <> dtLine Then modified = True
            
        divText.WriteLine dtLine
    Loop
                
    divText.Close
    Set divText = Nothing

    txtStrm.Close
    
    If modified Then
        verifyCSV = newFilePath
        opeLog.Add "氏名もしくはフリガナにスペースの混入を検知したため除去"
    Else
        verifyCSV = srcCSVpath
        On Error Resume Next
        Kill newFilePath
        On Error GoTo 0
    End If


End Function


Private Function removeSpace(ByVal dataline As String) As String
    Dim csvCells As Collection
    
    Set csvCells = splitCSVLine(dataline)


End Function


Private Function verifyCsvLine(ByVal titleLine As String, ByVal dataline As String) As String
    Dim csvCells As Collection
    
    Set csvCells = splitCSVLine(dataline)


End Function


Public Function chopString(ByVal tgtStr As String, ByVal byteLength As Long) As String
    Dim acIdx As Long
    Dim i As Long
    Dim bCnt As Long
    Dim tgtChr As String
    
    For i = 1 To Len(tgtStr)
        tgtChr = Mid(tgtStr, i, 1)
        acIdx = Asc(tgtChr)
        
        If acIdx >= 0 And acIdx <= 255 Then
            bCnt = bCnt + 1
        Else
            bCnt = bCnt + 2
        End If
        
        If bCnt <= byteLength Then
            chopString = chopString & tgtChr
        Else
            Exit For
        End If
    Next

End Function


Public Function moveFileAddHeadder(ByVal basePath As String, ByVal addString As String) As String
    Dim fso As FileSystemObject
    Dim fileName As String
    Dim fileExt As String
    Dim dirPath As String
    Dim retPath As String
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(basePath) Then Exit Function

    fileName = fso.GetBaseName(basePath)
    fileExt = fso.GetExtensionName(basePath)
    dirPath = fso.GetParentFolderName(basePath)
    
    Do
        retPath = fso.BuildPath(dirPath, addString & fileName & IIf(i = 0, vbNullString, "_" & i) & "." & fileExt)
        If Not fso.FileExists(retPath) Then Exit Do
        i = i + 1
    Loop
    
    On Error GoTo err
    fso.MoveFile basePath, retPath
    On Error GoTo 0
    
    moveFileAddHeadder = retPath

Exit Function
    
err:
    opeLog.Add "ファイル名を変更できませんでした。" & vbCrLf & "変更しようとしたファイル名：" & retPath


End Function


Public Function getDlFilePath(ByVal fileName As String, Optional ByVal msgFlg As Boolean = True) As String
    Dim wsh As Object 'WshShell
    Dim fso As Object 'FileSystemObject
    Dim folderPath As String
    Dim filePath As String
    
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    folderPath = SettingSh.DlFolderPath
    
    If InStr(folderPath, "ダウンロード") > 0 Then
        folderPath = Replace(folderPath, "ダウンロード", "Downloads")
    Else
        folderPath = folderPath
    End If
    
    folderPath = fso.GetAbsolutePathName(fso.BuildPath(wsh.SpecialFolders("MyDocuments") & "\..\", folderPath))
    filePath = fso.BuildPath(folderPath, fileName)
    
    If fso.FolderExists(folderPath) Then
        If fso.FileExists(filePath) Then
            getDlFilePath = filePath
        Else
            If msgFlg Then
                opeLog.Add filePath & vbCrLf & vbCrLf & "ダウンロード先に指定されたフォルダにファイルがありません。(上記パス)" & vbCrLf _
                    & "InternetExplorerの規定のダウンロードフォルダと、本ツールでダウンロード先に指定されたフォルダが一致しているか確認してください。" & vbCrLf & vbCrLf _
                    & "またダウンロードするファイルサイズが大きいために時間がかかっている場合は、タイムアウトする時間を延長してください。" & vbCrLf _
                    & "現在のタイムアウト設定（hh:mm:ss）：" & Format(SettingSh.DlTimeOut, "hh:mm:ss")
            End If
        End If
    Else
        If msgFlg Then
            opeLog.Add folderPath & vbCrLf & vbCrLf & "本ツールでダウンロード先に指定されたフォルダ（上記）がありません。"
        End If
    End If
    
    
End Function

Public Function putModifiedPeopele(ByVal argPeople As people, ByVal csvPath As String) As Boolean
    argPeople

End Function

Public Function getPeople(ByVal siteName As String, _
                            ByVal csvPath As String, _
                            Optional ByVal checkLabel As Boolean = True, _
                            Optional ByVal dateFrom As Date = 0) As people
                            
    Dim fso As Object 'new FileSystemObject
    Dim txtStrm As Object 'TextStream
    Dim dPeople As people

    If csvPath = vbNullString Then
        GoTo normalFin
    End If
    
    'On Error GoTo Err
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(csvPath) Then
        opeLog.Add csvPath & vbCrLf & vbCrLf & "ファイルが見つかりません。"
        Exit Function
    End If
    
    Set txtStrm = fso.OpenTextFile(csvPath)
    
    Set dPeople = New people
    
    Do While txtStrm.AtEndOfLine = False
        If txtStrm.Line = 1 Then
            If Not dPeople.setLabel(siteName, splitCSVLine(txtStrm.ReadLine), checkLabel) Then
                Exit Function
            End If
        Else
            If Not dPeople.setData(siteName, splitCSVLine(txtStrm.ReadLine), dateFrom) Then
                Exit Function
            End If
        End If
    Loop
    
    txtStrm.Close
    
    opeLog.Add dPeople.allPeople.Count & "人分のデータをロードしました。"

    'On Error GoTo 0
    
    Set getPeople = dPeople

normalFin:
    Set fso = Nothing
    Set txtStrm = Nothing
    
'    Exit Function
'Err:
'    MsgBox "CSVが空白か、もしくは読み取れません。" & vbCrLf _
'        & csvPath & vbCrLf _
'        & "上記ファイルを確認してください。", vbExclamation
'
'    'Set readCSVasDblCollection = Nothing
'    GoTo NormalFin
    
End Function

Private Function joinCSVLine(ByVal argCollection As Collection) As String
    Dim i As Long
    
    For i = 1 To argCollection.Count
        joinCSVLine = joinCSVLine & IIf(i = 1, "", ",") & convertCSVString(argCollection(i))
    Next

End Function

Private Function convertCSVString(ByVal cellValue As Variant, Optional ByVal dateFormat As String = vbNullString, Optional ByVal forceDblQuot As Boolean = False) As String
    Dim strCol As String
    
'    If forceDblQuot Then
'        strCol = """" & Replace(cellValue, """", """""") & """"
'    End If

    Select Case True
        '数値、日付への処理を優先すると「1,2,3」等のデータがエスケープできないので文字列に対する処理を優先
        Case InStr(cellValue, """"), forceDblQuot
            strCol = """" & Replace(cellValue, """", """""") & """"
        Case InStr(cellValue, ","), InStr(cellValue, vbCrLf), InStr(cellValue, vbLf)
            strCol = """" & cellValue & """"
        Case IsNumeric(cellValue)
            If cellValue = 0 Then
                strCol = "0" 'vbNullString
            Else
                strCol = CStr(CDbl(cellValue))
            End If
        Case IsDate(cellValue)
        'ハイフンは日付以外の可能性があるので除外
            If dateFormat = vbNullString Then
                strCol = cellValue
            ElseIf InStr(cellValue, "-") Then
                strCol = cellValue
            Else
                strCol = Format(cellValue, dateFormat)
            End If
        Case Else
            strCol = cellValue
    End Select
    
    convertCSVString = strCol

End Function

Public Sub trimRange(ByVal tgtRng As Range)
    Dim data As Variant
    Dim i As Long, j As Long
    
    data = tgtRng.Value
    
    For i = LBound(data, 1) To UBound(data, 1)
        For j = LBound(data, 2) To UBound(data, 2)
            data(i, j) = Trim(data(i, j))
        Next
    Next
    
    tgtRng.Value = data
    
End Sub


'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：getCurrentRegion
'//  機能　　　　　　：指定されたRangeから、タイトル行を除いたCurrentRegionを返す。
'//  引数　　　　　　：baseCellRange：起点となる範囲
'//                    titleRowSize: タイトル行の行数､オプション､デフォルトは0
'//  戻り値　　　　　：タイトル行を除いたCurrentRegion、取得できなければNothingを返す。
'//  作成者　　　　　：Akira Hashimoto
'//  作成日　　　　　：2017/12/28
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------

Function getCurrentRegion(ByVal baseCellRange As Range, _
                          Optional ByVal titleRowSize As Long = 0, _
                          Optional ByVal needAlart As Boolean = True) As Range
    
    'エラーが出た場合は、「CurError」に飛ぶ
    On Error GoTo CurError
                         
    '引数表開始セル(baseCellRange)、引数タイトル行数(titleRowSize)を基に、データを取得
    '【要変更】ResizeとOffset入れ替えにする
    With baseCellRange.CurrentRegion
        Set getCurrentRegion = .offset(titleRowSize, 0).Resize(.Rows.Count - titleRowSize, .Columns.Count)
    End With
    
    'エラーが出る場合　=　値がない場合
    If getCurrentRegion Is Nothing Then
        err.Raise Number:=GET_CUR_REG_ERR, Description:="Can not get Current Region"
    End If
    
    '値がある場合、エラーを無効にしてプロシージャを抜ける
    On Error GoTo 0
    Exit Function

'エラーが出た場合の処理
CurError:
    'Rangeオブジェクトに「Nothing」を代入
    Set getCurrentRegion = Nothing
    
    '引数(needAlart)がTrueの場合、アラートを表示
    If needAlart Then
        opeLog.Add baseCellRange.Parent.name & vbCrLf & "上記ワークシートのデータ範囲が取得が出来ませんでした。"
    End If
    
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：splitCSVLine
'//  機能　　　　　　：CSVの１行分をコンマで分割してコレクションで返します。基本的にはExcelで読んだのと同じ結果になる様しています。
'//  　　　　　　　　　注１:行の先頭、およびコンマの後のダブルクォーテーションのみフィールドのエスケープと解釈します｡
'//  　　　　　　　　　注２:フィールドがエスケイプされていないとき、ダブルクォーテーションでエスケイプしていないダブルクォーテーションも文字と解釈します｡
'//  引数　　　　　　：csvLine：CSVの１行分
'//  戻り値　　　　　：CSVのフィールドがitemであるCollection
'//  作成者　　　　　：Akira Hashimoto
'//  作成日　　　　　：2018/03/06
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------

'コレクションで取得
Function splitCSVLine(ByVal csvLine As String) As Collection
    Dim inDbQuot As Boolean
    Dim colValue As String
    Dim flgEsc As Boolean
    Dim char As String
    Dim i As Long
    
    Set splitCSVLine = New Collection
    
    If csvLine = vbNullString Then
        Exit Function
    End If
    
    For i = 1 To Len(csvLine)
        char = Mid(csvLine, i, 1)
    
        Select Case char
            Case """"
                If inDbQuot Then
                    If flgEsc Then
                        colValue = colValue & char
                    End If
                
                    flgEsc = Not flgEsc
                Else
                    If i = 1 Then
                        inDbQuot = True
                    Else
                        If Mid(csvLine, i - 1, 1) = "," Then
                            inDbQuot = True
                        Else
                            colValue = colValue & char
                        End If
                    End If
                End If
            
            Case ","
                If flgEsc Then
                    inDbQuot = False
                    flgEsc = False
                End If
                
                If inDbQuot Then
                    colValue = colValue & char
                Else
                    splitCSVLine.Add colValue
                    colValue = vbNullString
                End If
            Case Else
                If flgEsc Then
                    inDbQuot = False
                    flgEsc = False
                End If
            
                colValue = colValue & char
            
        End Select
    Next
    
    '最終カラム出力
    splitCSVLine.Add colValue

End Function

'Sub test()
'    makeDiffFile "C:\Users\武田圭\Downloads\90017スバル_マイ_セミナ0425152457.txt", "C:\Users\武田圭\Downloads\90017スバル_マイ_セミナ0424085108.txt", #4/26/2019 10:00:00 AM#
'
'End Sub


'差分ファイルを出力し、lastUpdateより古い日付をlastUpdateへ書き換える
Public Function makeDiffFile(ByVal newCSVPath As String, ByVal oldCSVpath As String, ByVal lastUpdate As Date) As String
    Dim outDirPath As String
    Dim olds As Collection
    Dim data As Collection
    Dim i As Long, j As Long
    
    Const TITLE_ROW As Long = 1
    
    '#CSV　→　Collection化
    Set data = getData(newCSVPath, 0)
    Set olds = getData(oldCSVpath, 0)
    
    If data Is Nothing Then GoTo abnormalFin
        
    '#同一データを削除（更新されたデータを残す）
    '#OLDが無い場合は飛ばす
    
    i = TITLE_ROW + 1
    Do Until data.Count < i Or olds Is Nothing
    
        j = TITLE_ROW + 1
        Do Until olds.Count < j
        
            If data(i) = olds(j) Then
                data.Remove i
                olds.Remove j
                Exit Do
            End If
            j = j + 1
        Loop
        
        If j = olds.Count + 1 Then i = i + 1
    Loop

    '#以下データ処理
    Dim targetCells As Collection
    Dim oldCells As Collection
    Dim cancelColumn As Long
    Dim updateColumn As Long
    Dim uniqTitles As Dictionary 'Object
    
    Dim targetCellValue As Variant
    Dim outLine As String
    Dim fso As Object 'new FileSystemObject
    Dim txtStrm As Object 'TextStream
    Dim outPath As String
    
    '#キャンセル列のタイトル、キャンセルを示す文字を設定

    Const CANCEL_TITLE = "キャンセルフラグ"
    Const CANCEL_TEXT = "1"
    Const UPDATE_TITLE = "エントリー日時"
     
    'データを特定するための列の組み合わせ。key = 列タイトル、item = 列番号
    Set uniqTitles = CreateObject("Scripting.Dictionary")
    
    uniqTitles.Add "学生管理ID", 0
    uniqTitles.Add "セミナー番号", 0
    uniqTitles.Add "エントリー日時", 0

    'タイトル行がなければエラー
    If TITLE_ROW < 1 Then
        MsgBox "タイトル行が必要です。", vbExclamation
        GoTo abnormalFin
    End If
    
    Set targetCells = splitCSVLine(data(TITLE_ROW))
    
    '最新ファイルのタイトルから「キャンセル」と「エントリー日時」、の列番号を探す
    For j = 1 To targetCells.Count
        targetCellValue = Trim(targetCells(j))
    
        If targetCellValue = CANCEL_TITLE Then
            cancelColumn = j
        ElseIf targetCellValue = UPDATE_TITLE Then
            updateColumn = j
        End If
        
        If uniqTitles.Exists(targetCellValue) Then
            uniqTitles(targetCellValue) = j
        End If
    Next
    
    'みつからなければエラー
    If cancelColumn < 1 Then
        MsgBox "最新のダウンロードデータに『" & CANCEL_TITLE & "』列がありません。", vbExclamation
        GoTo abnormalFin
    End If
    
    If updateColumn < 1 Then
        opeLog.Add "最新のダウンロードデータに『" & UPDATE_TITLE & "』列がありません。"
        GoTo abnormalFin
    End If
    
    For Each targetCellValue In uniqTitles
        If uniqTitles(targetCellValue) = 0 Then
            opeLog.Add "最新のダウンロードデータに『" & targetCellValue & "』列がありません。"
            GoTo abnormalFin
        End If
    Next
      
    ' ファイル出力の準備　パス作成とテキスト開く
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With fso
    
    i = 0
    Do
        outPath = .BuildPath(.GetParentFolderName(newCSVPath), .GetBaseName(newCSVPath) & _
                                "_差分" & IIf(i = 0, vbNullString, i) & ".csv")
        i = i + 1
    Loop While .FileExists(outPath)
    
    End With
    
    Set txtStrm = fso.CreateTextFile(outPath)


    'タイトル行を書き出し
    txtStrm.WriteLine data(TITLE_ROW)
    
    Dim k As Long, l As Long
    Dim tgtIdx As Long
    Dim MatchFlg As Boolean
    Dim oldCnt As Long
    
    If olds Is Nothing Then
        oldCnt = 0
    Else
        oldCnt = olds.Count
    End If
    '差分データが無い場合はタイトルのみで以下スキップ
    'ある場合は、
    For i = TITLE_ROW + 1 To data.Count
        Set targetCells = splitCSVLine(data(i))
        '差分ない場合はマッチフラグ　False
        MatchFlg = False
        
        For k = TITLE_ROW + 1 To oldCnt
            Set oldCells = splitCSVLine(olds(k))
            MatchFlg = True
            
            'データ特定の組み合わせが一つでも違えば抜ける
            For Each targetCellValue In uniqTitles
                tgtIdx = CLng(uniqTitles(targetCellValue))
                
                If oldCells(tgtIdx) <> targetCells(tgtIdx) Then
                    MatchFlg = False
                    Exit For
                End If
            Next
                       
            'データ全てが一致（MatchFlgがTrue）で「oldCells」を元のデータとしてループを抜ける
            If MatchFlg Then
                Exit For
            End If
        Next
        
        Dim nologgedCancel As Boolean
               
'        'エントリー日時が前回更新日より古く、かつ前回⇒今回でキャンセル状態に変わった場合は、エントリー日時を前回更新日に書き換える。
'        'エントリー日時が前回更新日より新しい「予約」ログがあった場合は、上記キャンセルログは上書きされるが、
'        'その場合「予約」ログがある時点でそれが最新であるので、上書きでOK。
'        If Not MatchFlg Then
'            nologgedCancel = targetCells(updateColumn) < lastUpdate And targetCells(cancelColumn) = CANCEL_TEXT
'        Else
'            nologgedCancel = targetCells(updateColumn) < lastUpdate And targetCells(cancelColumn) = CANCEL_TEXT And oldCells(cancelColumn) <> CANCEL_TEXT
'        End If
        
        '（マイナビ）キャンセルはタイムスタンプが更新されない
        
        If olds Is Nothing Then
            '差分をとっていない場合は、どれが差分かわからない。
            'ので、前回更新日-1日以降のログを対象とする。そこから前のキャンセルは対象としない。（古いままの日付で処理される）
            nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT And CDate(targetCells(updateColumn)) >= DateAdd("d", -1, lastUpdate)
        
        Else
            '差分をとった場合、
            If MatchFlg Then
                '前回⇒今回でキャンセル状態に変わった場合にUnloggedキャンセル（前回キャンセルだった場合は対応無用）
                nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT And oldCells(cancelColumn) <> CANCEL_TEXT
            Else
                '前回のデータがない場合はそのままUnloggedキャンセル（最新の日付はわからない）
                nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT
            End If
        End If
        
        'データをCSVに戻す
        For j = 1 To targetCells.Count
            'キャンセルの場合は、タイムスタンプがあてにならないので、タイムスタンプを前回更新日-1日で上書き
            'データ受信が毎朝定期でされているので、最大1日遅れる場合がある。つまり差分として出てきたキャンセルは少なくとも前回更新日の1日前以降に変更されたものと見れる。
            'この日時はiWeb側の履歴をどこまで追うかに効いてくる。
            If j = updateColumn And nologgedCancel Then
                If CDate(targetCells(updateColumn)) >= DateAdd("d", -1, lastUpdate) Then
                    outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & targetCells(j)
                Else
                    outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & Format(DateAdd("d", -1, lastUpdate), "yyyy/m/d hh:mm:ss")
                End If
            'キャンセル欄を 1 ⇒ 99　へ変更する。（ログが更新されていないキャンセルと、リクナビの通常のキャンセルを見分けるため）
            ElseIf j = cancelColumn And nologgedCancel Then
                outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & bookState.UnloggedCancel
            Else
                outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & targetCells(j)
            End If
        Next
            
        txtStrm.WriteLine outLine
        outLine = vbNullString
    Next
    
    txtStrm.Close
    makeDiffFile = outPath

normalFin:

Exit Function

abnormalFin:
    makeDiffFile = vbNullString
    
End Function

Private Function getData(ByVal csvPath As String, Optional ByVal titleRow As Long) As Collection
    Dim fso As Object 'new FileSystemObject
    Dim txtStrm As Object 'TextStream
    Dim i As Long

    If csvPath = vbNullString Then
        GoTo normalFin
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(csvPath) Then
        opeLog.Add csvPath & vbCrLf & "上記ファイルは見つかりません。"
        Exit Function
    End If
    
    Set txtStrm = fso.OpenTextFile(csvPath)
    
    Set getData = New Collection
    
    Do While txtStrm.AtEndOfLine = False
        If txtStrm.Line > titleRow Then
            getData.Add txtStrm.ReadLine
        Else
            txtStrm.SkipLine
        End If
    Loop
    
    txtStrm.Close

normalFin:
    Set fso = Nothing
    Set txtStrm = Nothing
    
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：getFilePathByDialog
'//  機能　　　　　　：ダイアログを使用してユーザーに単一のファイルを選ばせ、そのパスを返す。
'//  引数　　　　　　：ダイアログのフィルター（ディスクリプション、拡張子）とタイトル。
'//  　　　　　　　　　いずれもオプション｡
'//  戻り値　　　　　：フォルダのパス/vbNullString
'//  作成者　　　　　：Akira Hashimoto
'//  作成日　　　　　：2017/12/11
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------

Public Function getFilePathByDialog(Optional ByVal argExt As String = "*.*", _
                            Optional ByVal argDscr As String = "All files", _
                            Optional ByVal argTitle As String = "ファイルを選択して下さい") As String
    Const FILE_PICKER = 3

    With Application.FileDialog(FILE_PICKER)
        .Title = argTitle
        .Filters.Clear
        .Filters.Add argDscr, argExt
        .InitialFileName = ""
        .AllowMultiSelect = False
        
        If .Show = True Then
            getFilePathByDialog = .SelectedItems(1)
        Else
            getFilePathByDialog = vbNullString
        End If
    End With

End Function
