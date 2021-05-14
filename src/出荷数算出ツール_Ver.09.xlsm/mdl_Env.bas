Attribute VB_Name = "mdl_Env"
Option Explicit
Option Private Module


'============================================================
'　ファイル周り・ファイル周り
'============================================================

'ファイルダイアログを開き、ツールシートのファイル設定欄にユーザが選択したファイルパスを入れる
Sub SubSetCSVFilePathByUserChoice(ByVal uexlSpecify As ueXLFileType)

    Const INITIAL_FILE    As String = "C:\"

    Dim myRange         As Range
    Dim strFolderPath   As String
    Dim strFilePath     As String

    'ファイルダイアログが開くので実施確認は出さない
    Call SubCtrlMovableCmd(ueppStart)

    'プリセットとしてファイル設定欄のセルを取得(Nothing返却はない)
    Set myRange = FncGetRangeOfFileSetting(uexlSpecify) '書く時に使うのでRangeを取る
    If myRange.Value <> "" Then
        'ファイルのフォルダを取得
        strFilePath = FncGetFolderNameFromFilePath(myRange.Value)
        strFolderPath = FncChkPathEnd(strFilePath)

        'シート欄が空欄の場合このファイルのフォルダに
    Else
        strFolderPath = ThisWorkbook.Path
    End If


    '初期値渡しながらダイアログ開く
    strFilePath = FncOpenDialogAndGetFile(strFolderPath)

    If strFilePath = "" Then                     '空欄はキャンセル
        Call Subキャンセル表示
    Else

        Application.ScreenUpdating = True        '即見えるように
        myRange.Value = strFilePath

        Call Sub正常終了表示(, "ファイルを設定しました")

    End If

    Set myRange = Nothing

End Sub

'ファイル選択処理(キャンセルやエラーは空白が返る)
Function FncOpenDialogAndGetFile(Optional ByVal strInitialPath As String = "") As String

    Const DEF_TITLE     As String = "CSVファイルの選択"

    Dim strFilterName   As String
    Dim strFilterExt    As String
    Dim strGetResult    As String

    strFilterName = G_FILTERNAME_EXCEL
    strFilterExt = "*" & G_EXT_XLS

    '実行
    With Application.FileDialog(msoFileDialogFilePicker)

        '複数選択の可不可(このプロシージャでは不可限定)
        .AllowMultiSelect = False

        'フィルタのクリア
        .Filters.Clear

        'ファイルフィルタの追加
        .Filters.Add strFilterName, strFilterExt

        '初期表示ファイルの設定
        If strInitialPath <> "" Then
            .InitialFileName = strInitialPath
        End If

        'ダイアログタイトル
        .Title = DEF_TITLE

        '結果取得
        If .Show = True Then
            strGetResult = .SelectedItems(1)
        Else
            strGetResult = vbNullString          '=""
        End If

    End With

    FncOpenDialogAndGetFile = strGetResult


End Function

'ファイル設定欄のセルを返す(Nothing返却はない)
Function FncGetRangeOfFileSetting(ByVal uexlSpecify As ueXLFileType) As Range

    Const FOOTER_NAME       As String = "ファイル名"
    Const SAFE_ADDRESS_OA   As String = "D7"     '検収_入荷
    Const SAFE_ADDRESS_SS   As String = "D12"    '商品台帳（期末）
    Const SAFE_ADDRESS_DD   As String = "D17"    '商品台帳（期首）
    Const SAFE_ADDRESS_SU   As String = "D30"    '仕入先

    Dim myRange     As Range
    Dim strTarget   As String
    Dim strAddress  As String

    Select Case uexlSpecify
        Case uexlOrderArrival
            strTarget = G_FILE_TARGET_OA & FOOTER_NAME
            strAddress = SAFE_ADDRESS_OA
        Case uexlPreProductBook
            strTarget = G_FILE_TARGET_SS & FOOTER_NAME
            strAddress = SAFE_ADDRESS_SS
        Case uexlProductBook
            strTarget = G_FILE_TARGET_DD & FOOTER_NAME
            strAddress = SAFE_ADDRESS_DD
        Case uexlSupplierBook
            strTarget = G_FILE_TARGET_SU & FOOTER_NAME
            strAddress = SAFE_ADDRESS_SU
    End Select

    'ファイル設定欄を検索
    Set myRange = FncGetTagetRange(strTarget, G_SHEETNAME_TOOL, 0, 1)

    If myRange Is Nothing Then
        Set myRange = ThisWorkbook.Worksheets(G_SHEETNAME_TOOL).Range(strAddress)
    End If

    Set FncGetRangeOfFileSetting = myRange
    Set myRange = Nothing

End Function

'ツールシートのファイル設定を取得してファイル有無がTrueならファイルパスを返す
Function FncGetFileSetting(ByVal uexlSpecify As ueXLFileType) As String

    '定数
    Const ERR_MESSAGE_N1    As String = "ファイルが設定されていません" & vbCrLf
    Const ERR_MESSAGE_N2    As String = "ファイルを設定してから再実行して下さい"
    Const ERR_MESSAGE_N3    As String = "安全在庫数は出ませんが"
    Const ERR_MESSAGE_C1    As String = "指定のファイルが存在しません" & vbCrLf
    Const ERR_MESSAGE_C2    As String = "ファイル設定かフォルダ内を確認して" & vbCrLf & _
    "正しいファイル名を設定して下さい"

    '変数
    Dim myRange             As Range
    Dim strFilePath         As String
    Dim strErrorMessage     As String
    Dim lngErrNum           As Long
    Dim blnResult           As Boolean

    'ファイル設定欄のセルを取得(Nothing返却はない)
    Set myRange = FncGetRangeOfFileSetting(uexlSpecify)
    
    Dim TargetFile As String
    Select Case uexlSpecify
        Case uexlOrderArrival
            TargetFile = G_FILE_TARGET_OA
        Case uexlPreProductBook
            TargetFile = G_FILE_TARGET_SS
        Case uexlProductBook
            TargetFile = G_FILE_TARGET_DD
        Case uexlSupplierBook
            TargetFile = G_FILE_TARGET_SU
    End Select
    '欄が空だったら
    If myRange.Value = "" Then
        lngErrNum = G_CTRL_ERROR_NUMBER_USER_CAUTION
        strErrorMessage = TargetFile & ERR_MESSAGE_N1 & ERR_MESSAGE_N2
        GoTo endRoutine
    End If
    strFilePath = myRange.Value

    '指定ファイルの有無確認
    blnResult = FncCheckFileExist(strFilePath)
    If blnResult = False Then
        lngErrNum = G_CTRL_ERROR_NUMBER_USER_CAUTION
        strErrorMessage = TargetFile & ERR_MESSAGE_C1 & ERR_MESSAGE_C2
        GoTo endRoutine
    End If

    FncGetFileSetting = strFilePath


endRoutine:
    Set myRange = Nothing
    If lngErrNum <> 0 Then                       'Err.Numberとは違うので注意
        Err.Raise lngErrNum, , strErrorMessage
    End If

End Function

Public Sub GetFoldePath()
    On Error GoTo ErrHdl
    ThisWorkbook.Worksheets("元データ読込").Range(SAVE_FOLDER).Value = ShowFileDialog

ExitHdl:
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
    Resume ExitHdl
End Sub

Private Function ShowFileDialog() As String
    Dim temp As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then
            temp = .SelectedItems(1)
        End If
    End With
    ShowFileDialog = temp
End Function


