Attribute VB_Name = "mdl_Common"
Option Explicit
Option Private Module


'============================================================
'　共通機能・汎用処理
'============================================================

'----------------
'　列挙型
'----------------
'処理の開始時か終了時か
Public Enum uePrePost
    [_MIN] = -1
    ueppEnd = 0
    ueppStart = 1
    [_MAX]
End Enum

'1列目か2列目か
Public Enum ueColPos
    uecpCol1
    uecpCol2
End Enum

Public Const SAVE_FOLDER As String = "D22"
'----------------
'　メソッド
'----------------

'実行確認及びOKなら前準備
Function Fnc実行前確認(Optional strTitle As String = "", Optional strAddMessage As String = "") As Boolean

    Dim blnResult       As Boolean

    blnResult = Fnc実行確認表示(strTitle, strAddMessage)

    If blnResult = False Then
        Call Subキャンセル表示(strTitle)
        Call SubCtrlMovableCmd(ueppEnd)          '念のため画面制御Off
        blnResult = False
    Else
        '画面制御を入れる
        Call SubCtrlMovableCmd(ueppStart)
        blnResult = True
    End If

    Fnc実行前確認 = blnResult

End Function

'実行確認表示、ユーザ選択回答をBoolean型で返す
Function Fnc実行確認表示(Optional strTitle As String = "", Optional strAddMessage As String = "") As Boolean

    Const DEFAULT_MESSAGE   As String = "処理を実行しますか？"
    Dim strMessage      As String
    Dim blnResult       As Boolean
    Dim vbmbrResult     As VbMsgBoxResult

    If strTitle = "" Then
        strTitle = "確認"
    End If

    If strAddMessage <> "" Then
        strAddMessage = strAddMessage & vbCrLf
    End If

    strMessage = strAddMessage & DEFAULT_MESSAGE

    vbmbrResult = MsgBox(strMessage, vbQuestion + vbOKCancel, strTitle)
    If vbmbrResult = vbOK Then
        blnResult = True
    ElseIf vbmbrResult = vbCancel Then
        blnResult = False
    End If

    Fnc実行確認表示 = blnResult

End Function

'キャンセル表示
Sub Subキャンセル表示(Optional strTitle As String = "", Optional strAddMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "キャンセルされました"
    Dim strMessage      As String
    Dim vbmbrResult     As VbMsgBoxResult


    '画面制御を解放 (軽くなったので先に開放)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "処理中止"
    Else
        strTitle = strTitle & "処理中止"
    End If

    strMessage = strAddMessage & DEFAULT_MESSAGE

    vbmbrResult = MsgBox(strMessage, vbInformation + vbOKOnly, strTitle)

    '    '重いので先にメッセージ
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'正常終了アナウンス
Sub Sub正常終了表示(Optional strTitle As String = "", Optional strMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "処理が正常に終了しました"
    Dim vbmbrResult     As VbMsgBoxResult


    '画面制御を解放 (軽くなったので先に開放)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "正常終了"
    End If

    If strMessage = "" Then
        strMessage = DEFAULT_MESSAGE
    End If

    vbmbrResult = MsgBox(strMessage, vbInformation + vbOKOnly, strTitle)

    '    '重いので先にメッセージ
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'ワーニング終了表示（判定でエラーにした時などに使う）
Sub Subワーニング終了表示(Optional strTitle As String = "", Optional strMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "異常終了" & vbCrLf & vbCrLf
    Dim vbmbrResult     As VbMsgBoxResult


    '画面制御を解放 (軽くなったので先に開放)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "エラー"
    End If

    If strMessage = "" Then
        strMessage = DEFAULT_MESSAGE
    End If

    vbmbrResult = MsgBox(strMessage, vbExclamation + vbOKOnly, strTitle)

    '    '重いので先にメッセージ
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'システムエラー発生
Sub Subシステムエラー発生(Optional strTitle As String = "", Optional strAddMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "システムエラーが発生しました" & vbCrLf & vbCrLf
    Dim strMessage      As String
    Dim vbmbrResult     As VbMsgBoxResult


    '画面制御を解放 (軽くなったので先に開放)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "システムエラー発生"
    End If

    strMessage = DEFAULT_MESSAGE & strAddMessage

    vbmbrResult = MsgBox(strMessage, vbCritical + vbOKOnly, strTitle)

    '    '重いので先にメッセージ
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'エラー制御
Sub SubHandleErrorAndFinishing(Optional ByVal strTitle As String = "")

    Select Case Err.Number
        Case 0
            Call Sub正常終了表示(strTitle)             '正常終了もまとめたい時

        Case G_CTRL_ERROR_NUMBER_USER_NOTICE     'エラーではないけど注釈
            Call Sub正常終了表示(strTitle, Err.Description)

        Case G_CTRL_ERROR_NUMBER_USER_CAUTION    'ユーザーエラー
            Call Subワーニング終了表示(strTitle, Err.Description)

        Case G_CTRL_ERROR_NUMBER_DEVELOPER       '開発者エラー
            Call Subワーニング終了表示(strTitle, Err.Description)

        Case Else                                'システムエラー
            Call Subシステムエラー発生(strTitle, Err.Description)

    End Select

End Sub

'渡されたフォルダパスの最後が￥でなかったら付けて返す
Function FncChkPathEnd(ByVal strFolderPath As String) As String

    Const FOLDER_MARK   As String = "\"

    If Right(strFolderPath, 1) <> FOLDER_MARK Then
        strFolderPath = strFolderPath & FOLDER_MARK
    End If
    FncChkPathEnd = strFolderPath

End Function

'渡されたファイルパスのフォルダパスを返す
Function FncGetFolderNameFromFilePath(ByVal strFilePath As String) As String

    Dim myFSO       As FileSystemObject
    Dim strFolder   As String

    Set myFSO = New FileSystemObject

    strFolder = myFSO.GetParentFolderName(strFilePath)

    FncGetFolderNameFromFilePath = strFolder

    Set myFSO = Nothing

End Function

'渡されたフォルダパスが存在するかチェックして結果を返す
Function FncCheckFolderExist(ByVal strFolderPath As String) As Boolean

    Dim myFSO       As FileSystemObject
    Dim blnResult   As Boolean

    Set myFSO = New FileSystemObject

    blnResult = myFSO.FolderExists(strFolderPath)

    FncCheckFolderExist = blnResult

    Set myFSO = Nothing

End Function

'渡されたファイルパスが存在するかチェックして結果を返す
Function FncCheckFileExist(ByVal strFilePath As String) As Boolean

    Dim myFSO       As FileSystemObject
    Dim blnResult   As Boolean

    Set myFSO = New FileSystemObject

    blnResult = myFSO.FileExists(strFilePath)

    FncCheckFileExist = blnResult

    Set myFSO = Nothing

End Function

'ファイルコピー(上書き型)
Function FncCopySpecifFile(ByVal strCopyPath As String, ByVal strPastePath As String) As Boolean

    Dim myFSO       As FileSystemObject

    Set myFSO = New FileSystemObject

    myFSO.CopyFile strCopyPath, strPastePath, True

    FncCopySpecifFile = True                     'システムエラーが生じた時点でここは通らない

    Set myFSO = Nothing

End Function

'ファイルを削除
Function FncDeleteFile(ByVal strFilePath As String) As Boolean

    Dim myFSO       As FileSystemObject

    Set myFSO = New FileSystemObject

    myFSO.DeleteFile strFilePath, True

    FncDeleteFile = True                         'システムエラーが生じた時点でここは通らない

    Set myFSO = Nothing

End Function

'テキスト形式でファイルを読み込み、シートを返す
Function FncOpenTextFileLegacy(ByVal strFilePath As String, _
Optional ByVal vntFieldInfo As Variant) As Worksheet

    Dim WS      As Worksheet

    With Workbooks
        .OpenText Filename:=strFilePath, _
        Origin:=932, StartRow:=1, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, _
        Tab:=True, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
        FieldInfo:=vntFieldInfo, _
        TrailingMinusNumbers:=True

        Set WS = Workbooks(.Count).Worksheets(1)
    End With

    Set FncOpenTextFileLegacy = WS
    Set WS = Nothing

End Function

'指定列でソートをかける(1行目項目名、昇順限定)
Function FncSortWithSpecifColumn(ByVal rngSortArea As Range, ByRef alngKeyColumn() As Long) As Boolean

    Dim WS      As Worksheet
    Dim c       As Long

    Set WS = rngSortArea.Parent

    With WS.Sort
        With .SortFields
            .Clear
            For c = LBound(alngKeyColumn) To UBound(alngKeyColumn)
                .Add KEY:=rngSortArea.Cells(1, alngKeyColumn(c)), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortTextAsNumbers
            Next c
        End With

        .SetRange rngSortArea
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    '痕跡消去
    With rngSortArea.Cells(1, 1)
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With

    FncSortWithSpecifColumn = True
    Set WS = Nothing
    Set rngSortArea = Nothing

End Function

'セル範囲から指定の文字列のセルを返す（該当なし＝Nothingが返る）
Function FncGetTagetRangeFromRange(ByVal strTarget As String, ByVal TargetRange As Range, _
Optional ByVal lngRowOfs As Long, Optional ByVal lngColOfs As Long) _
As Range

    Dim myRange     As Range
    Dim lngRows     As Long
    Dim lngCols     As Long
    Dim r           As Long
    Dim c           As Long

    For r = 1 To TargetRange.Rows.Count
        For c = 1 To TargetRange.Columns.Count
            If TargetRange(r, c) = strTarget Then
                lngRows = r
                lngCols = c
                Exit For
            End If
        Next c
        If lngCols <> 0 Then Exit For
    Next r

    If lngRows <> 0 And lngCols <> 0 Then
        Set myRange = TargetRange.Cells(lngRows, lngCols).Offset(lngRowOfs, lngColOfs)
    End If

    Set FncGetTagetRangeFromRange = myRange
    Set myRange = Nothing
    Set TargetRange = Nothing

End Function

'指定の文字列のセルを返す（該当なし＝Nothingが返る）
Function FncGetTagetRange(ByVal strTarget As String, ByVal strSheetName As String, _
Optional ByVal lngRowOfs As Long, Optional ByVal lngColOfs As Long) _
As Range

    Dim myRange     As Range
    Dim vntRange    As Variant
    Dim lngRows     As Long
    Dim lngCols     As Long
    Dim r           As Long
    Dim c           As Long

    With ThisWorkbook.Worksheets(strSheetName)
        vntRange = Range(.Cells(1, 1), .Cells.SpecialCells(xlCellTypeLastCell)).Value
    End With

    For r = 1 To UBound(vntRange, ueRC.uercRow)
        For c = 1 To UBound(vntRange, ueRC.uercCol)
            If vntRange(r, c) = strTarget Then
                lngRows = r
                lngCols = c
                Exit For
            End If
        Next c
        If lngCols <> 0 Then Exit For
    Next r

    If lngRows <> 0 And lngCols <> 0 Then
        With ThisWorkbook.Worksheets(strSheetName)
            Set myRange = .Cells(lngRows, lngCols).Offset(lngRowOfs, lngColOfs)
        End With
    End If

    Set FncGetTagetRange = myRange
    Set myRange = Nothing
    If IsArray(vntRange) Then Erase vntRange

End Function

'指定の文字列のセルを返すFind版（該当なし＝Nothingが返る）
Function FncFindTagetRange(ByVal strTarget As String, ByVal strSheetName As String, _
Optional ByVal lngRowOfs As Long, Optional ByVal lngColOfs As Long) _
As Range

    Dim myRange     As Range

    With ThisWorkbook.Worksheets(strSheetName)

        Set myRange = .Cells.Find(What:=strTarget, After:=.Cells(1, 1), LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext) _
        .Offset(lngRowOfs, lngColOfs)

    End With

    Set FncFindTagetRange = myRange
    Set myRange = Nothing

End Function

'2次元配列をキー位置でソートする (1次元と2次元限定)
Function FncBubbleSortFor2DArray(ByVal avntArray As Variant, ByVal uecpKeyPos As ueColPos) As Variant()

    Dim avnt2DArray()   As Variant
    Dim vntSwap     As Variant
    Dim lngRowMin   As Long, lngRowMax   As Long
    Dim lngColMin   As Long, lngColMax   As Long
    Dim lngRank     As Long
    Dim r           As Long
    Dim r2          As Long
    Dim c           As Long


    lngRank = FncGetArrayDimension(avntArray)
    If lngRank = 0 Then
        Err.Raise G_CTRL_ERROR_NUMBER_DEVELOPER, , "配列以外使用不可"
    ElseIf lngRank > 2 Then
        Err.Raise G_CTRL_ERROR_NUMBER_DEVELOPER, , "1次元配列か2次元配列のみ使用可"
    End If

    lngRowMin = LBound(avntArray, ueRC.uercRow): lngRowMax = UBound(avntArray, ueRC.uercRow)

    If lngRowMax - lngRowMin = 0 Then GoTo endRoutine

    If lngRank = 1 Then
        lngColMin = 1: lngColMax = 2
        ReDim avnt2DArray(lngColMin To lngColMax, 1 To 1)
        For r = lngRowMin To lngRowMax
            avnt2DArray(r, 1) = avntArray(r)
        Next r
        avntArray = avnt2DArray
    ElseIf lngRank = 2 Then
        lngColMin = LBound(avntArray, ueRC.uercCol): lngColMax = UBound(avntArray, ueRC.uercCol)
    End If

    For r = lngRowMin To lngRowMax
        For r2 = lngRowMin To lngRowMax - 1
            If avntArray(r2, uecpKeyPos) > avntArray(r2 + 1, uecpKeyPos) Then
                For c = lngColMin To lngColMax
                    vntSwap = avntArray(r2, c)
                    avntArray(r2, c) = avntArray(r2 + 1, c)
                    avntArray(r2 + 1, c) = vntSwap
                Next
            End If
        Next r2
    Next r

endRoutine:
    FncBubbleSortFor2DArray = avntArray
    Erase avntArray
    Erase avnt2DArray

End Function

'指定ピボット名のあるシート名を返す
Function FncGetSheetNameWithinPivot(ByVal strPivotName As String) As String

    Dim WS      As Worksheet
    Dim PT      As PivotTable
    Dim strName As String

    For Each WS In ThisWorkbook.Worksheets
        For Each PT In WS.PivotTables
            If PT.Name = strPivotName Then
                strName = WS.Name
                GoTo endRoutine
            End If
        Next PT
    Next WS

endRoutine:
    FncGetSheetNameWithinPivot = strName
    Set PT = Nothing
    Set WS = Nothing

End Function

'指定ピボットの範囲変更（2つまでソート指定可）
Function FncChangePivotSourceArea(ByVal myRange As Range, _
Optional ByVal strPivotName As String = "", _
Optional ByVal strSheetName As String = "", _
Optional ByVal strFieldName1 As String = "", _
Optional ByVal strFieldName2 As String = "") As Boolean

    Dim WS          As Worksheet
    Dim myPivot     As PivotTable
    Dim PF1         As PivotField
    Dim PF2         As PivotField

    On Error GoTo endRoutine

    If strSheetName = "" Then
        Set WS = myRange.Parent
    Else
        Set WS = ThisWorkbook.Worksheets(strSheetName)
    End If

    If strPivotName = "" Then
        Set myPivot = WS.PivotTables(1)
    Else
        Set myPivot = WS.PivotTables(strPivotName)
    End If

    With myPivot                                 'External:=True必須
        .SourceData = myRange.Address(True, True, ReferenceStyle:=xlR1C1, External:=True)
        .RefreshTable

        If strFieldName1 <> "" Then
            Set PF1 = .PivotFields(strFieldName1)
            PF1.AutoSort Order:=xlAscending, Field:=PF1.Name
        End If
        If strFieldName2 <> "" Then
            Set PF2 = .PivotFields(strFieldName2)
            PF2.AutoSort Order:=xlAscending, Field:=PF2.Name
        End If
        .RefreshTable
    End With

    FncChangePivotSourceArea = True

endRoutine:
    Set PF1 = Nothing
    Set PF2 = Nothing
    Set myPivot = Nothing
    Set myRange = Nothing
    Set WS = Nothing

End Function

'指定ピボットのあるセル範囲を返す（見つからなければエラー終了）
Function FncGetTableRangeOfPivot(ByVal strPivotName As String, _
ByVal strSheetName As String) As Range

    Dim WS          As Worksheet
    Dim myRange     As Range
    Dim myPivot     As PivotTable

    On Error GoTo endRoutine

    If strSheetName = "" Then
        strSheetName = FncGetSheetNameWithinPivot(strPivotName)
        If strSheetName = "" Then
            Set WS = ActiveSheet
        End If
    End If
    If WS Is Nothing Then
        Set WS = ThisWorkbook.Worksheets(strSheetName)
    End If

    Set myPivot = WS.PivotTables(strPivotName)
    If myPivot Is Nothing Then
        Err.Raise G_CTRL_ERROR_NUMBER_USER_CAUTION, , _
        "指定のピボットテーブルがありません" & vbCrLf & _
        "ピボットテーブルは削除されたか名前が変更された可能性があります"
    End If

    Set myRange = myPivot.TableRange1

endRoutine:
    Set FncGetTableRangeOfPivot = myRange
    Set myRange = Nothing
    Set myPivot = Nothing
    Set WS = Nothing

    If Err.Number <> 0 Then
        Call SubHandleErrorAndFinishing("ピボット位置探索")
    End If

End Function

'ピボットテーブルのセル範囲を渡し、有効状態（値が入っている）かどうか返す(Trueかエラーメッセージ付きFalse)
Function FncChekPivotRangeValid(ByVal rngPivot As Range, _
Optional ByVal blnRow As Boolean = True, _
Optional ByVal blnColumn As Boolean = True) As Boolean

    Const ERROR_NONE    As String = "ピボットテーブルに有効データがありません" & vbCrLf & _
    "データを読み込んでから再実行して下さい"
    Const ERROR_INVALID_ARGUMENT    As String = "RowとColumnの両方Falseにすることは出来ません"

    Const MIN_ROWSIZE   As Long = 4
    Const MIN_COLSIZE   As Long = 3

    Dim strMessage      As String
    Dim lngRowSize      As Long
    Dim lngColSize      As Long
    Dim blnRowResult    As Boolean
    Dim blnColResult    As Boolean
    Dim blnTotalResult  As Boolean

    If blnRow = False And blnColumn = False Then
        Err.Raise G_CTRL_ERROR_NUMBER_DEVELOPER, , ERROR_INVALID_ARGUMENT
    End If

    With rngPivot
        lngRowSize = .Rows.Count
        lngColSize = .Columns.Count
    End With

    If lngRowSize > MIN_ROWSIZE Then
        blnRowResult = True
    End If
    
    If lngColSize > MIN_COLSIZE Then
        blnColResult = True
    End If

    If blnRow And blnColumn Then
        blnTotalResult = blnRowResult And blnColResult
    Else
        If blnRow Then
            blnTotalResult = blnRowResult
        ElseIf blnRow Then
            blnTotalResult = blnColResult
        End If
    End If

    If blnTotalResult Then
        FncChekPivotRangeValid = blnTotalResult
    Else
        'Falseが返る
        Err.Raise G_CTRL_ERROR_NUMBER_USER_CAUTION, , ERROR_NONE
    End If

End Function

'配列の次元数を取得する
Function FncGetArrayDimension(ByVal avntArray As Variant) As Long

    Dim d       As Long
    Dim tmp     As Long

    If Not IsArray(avntArray) Then Exit Function '0が返る

    d = 1
    On Error Resume Next
    Do

        tmp = UBound(avntArray, d)
        If Err.Number = 0 Then
            d = d + 1
        Else
            d = d - 1
            Exit Do
        End If

    Loop While Err.Number = 0
    Err.Clear
    On Error GoTo 0

    FncGetArrayDimension = d
    Erase avntArray

End Function

'データが整数かどうか返す（投入は数値のみ）
Function IsInteger(ByVal vntData As Variant) As Boolean

    Dim blnResult       As Boolean

    If IsNumeric(vntData) Then
        If Int(vntData) = vntData Then
            blnResult = True
        End If
    End If

    IsInteger = blnResult

End Function

'渡されたデータが数字のみかどうかを返す(小数点、ハイフン、記号はダメ)
Function FncCheckDataIsNumberOnly(ByVal vntData As Variant) As Boolean

    Const NUM_LIST  As String = "1234567890"

    Dim lngLen      As Long
    Dim l           As Long
    Dim lngResult   As Long
    Dim blnResult   As Boolean

    If IsNumeric(vntData) Then
        lngLen = Len(vntData)
        For l = 1 To lngLen
            lngResult = InStr(NUM_LIST, Mid(vntData, l, 1))
            If lngResult = 0 Then
                Exit For
            End If
        Next l
    End If

    If lngResult <> 0 Then                       '数字じゃなければ0が入ってる
        blnResult = True
    End If

    FncCheckDataIsNumberOnly = blnResult

End Function

'配列の中にエラー値があるか
Function FncCheckErrorInArray(ByVal vntArray As Variant) As Boolean

    Dim vntCurrent      As Variant
    Dim blnResult       As Boolean

    If IsArray(vntArray) Then

        For Each vntCurrent In vntArray

            If IsError(vntCurrent) Then
                blnResult = True
                Exit For
            End If

        Next vntCurrent

    Else
        blnResult = IsError(vntArray)
    End If

    FncCheckErrorInArray = blnResult

End Function

'文字列を文字列配列に変換して返す
Function FncTransfarStringToArray(ByVal strItemLine As String) As String()

    Dim vntItems    As Variant
    Dim lngCols     As Long

    vntItems = Split(strItemLine, ",")
    lngCols = UBound(vntItems) + 1
    ReDim Preserve vntItems(1 To lngCols)

    FncTransfarStringToArray = vntItems
    If IsArray(vntItems) Then Erase vntItems

End Function

'
''クラス起動関数
'Function Initialize(ByVal clsType As IInitializeAtConstOnly, _
'                    ByRef OperationSpecify As Variant, _
'                    Optional ByRef OperationSpecify2 As Variant) As IInitializeAtConstOnly
'
'    If IsMissing(OperationSpecify2) Then
'        Set Initialize = clsType.IInit(OperationSpecify)
'    Else
'        Set Initialize = clsType.IInit(OperationSpecify, OperationSpecify2)
'    End If
'
'End Function


