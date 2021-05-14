Attribute VB_Name = "mdl_Common"
Option Explicit
Option Private Module


'============================================================
'�@���ʋ@�\�E�ėp����
'============================================================

'----------------
'�@�񋓌^
'----------------
'�����̊J�n�����I������
Public Enum uePrePost
    [_MIN] = -1
    ueppEnd = 0
    ueppStart = 1
    [_MAX]
End Enum

'1��ڂ�2��ڂ�
Public Enum ueColPos
    uecpCol1
    uecpCol2
End Enum

Public Const SAVE_FOLDER As String = "D22"
'----------------
'�@���\�b�h
'----------------

'���s�m�F�y��OK�Ȃ�O����
Function Fnc���s�O�m�F(Optional strTitle As String = "", Optional strAddMessage As String = "") As Boolean

    Dim blnResult       As Boolean

    blnResult = Fnc���s�m�F�\��(strTitle, strAddMessage)

    If blnResult = False Then
        Call Sub�L�����Z���\��(strTitle)
        Call SubCtrlMovableCmd(ueppEnd)          '�O�̂��߉�ʐ���Off
        blnResult = False
    Else
        '��ʐ��������
        Call SubCtrlMovableCmd(ueppStart)
        blnResult = True
    End If

    Fnc���s�O�m�F = blnResult

End Function

'���s�m�F�\���A���[�U�I���񓚂�Boolean�^�ŕԂ�
Function Fnc���s�m�F�\��(Optional strTitle As String = "", Optional strAddMessage As String = "") As Boolean

    Const DEFAULT_MESSAGE   As String = "���������s���܂����H"
    Dim strMessage      As String
    Dim blnResult       As Boolean
    Dim vbmbrResult     As VbMsgBoxResult

    If strTitle = "" Then
        strTitle = "�m�F"
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

    Fnc���s�m�F�\�� = blnResult

End Function

'�L�����Z���\��
Sub Sub�L�����Z���\��(Optional strTitle As String = "", Optional strAddMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "�L�����Z������܂���"
    Dim strMessage      As String
    Dim vbmbrResult     As VbMsgBoxResult


    '��ʐ������� (�y���Ȃ����̂Ő�ɊJ��)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "�������~"
    Else
        strTitle = strTitle & "�������~"
    End If

    strMessage = strAddMessage & DEFAULT_MESSAGE

    vbmbrResult = MsgBox(strMessage, vbInformation + vbOKOnly, strTitle)

    '    '�d���̂Ő�Ƀ��b�Z�[�W
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'����I���A�i�E���X
Sub Sub����I���\��(Optional strTitle As String = "", Optional strMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "����������ɏI�����܂���"
    Dim vbmbrResult     As VbMsgBoxResult


    '��ʐ������� (�y���Ȃ����̂Ő�ɊJ��)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "����I��"
    End If

    If strMessage = "" Then
        strMessage = DEFAULT_MESSAGE
    End If

    vbmbrResult = MsgBox(strMessage, vbInformation + vbOKOnly, strTitle)

    '    '�d���̂Ő�Ƀ��b�Z�[�W
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'���[�j���O�I���\���i����ŃG���[�ɂ������ȂǂɎg���j
Sub Sub���[�j���O�I���\��(Optional strTitle As String = "", Optional strMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "�ُ�I��" & vbCrLf & vbCrLf
    Dim vbmbrResult     As VbMsgBoxResult


    '��ʐ������� (�y���Ȃ����̂Ő�ɊJ��)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "�G���["
    End If

    If strMessage = "" Then
        strMessage = DEFAULT_MESSAGE
    End If

    vbmbrResult = MsgBox(strMessage, vbExclamation + vbOKOnly, strTitle)

    '    '�d���̂Ő�Ƀ��b�Z�[�W
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'�V�X�e���G���[����
Sub Sub�V�X�e���G���[����(Optional strTitle As String = "", Optional strAddMessage As String = "")

    Const DEFAULT_MESSAGE   As String = "�V�X�e���G���[���������܂���" & vbCrLf & vbCrLf
    Dim strMessage      As String
    Dim vbmbrResult     As VbMsgBoxResult


    '��ʐ������� (�y���Ȃ����̂Ő�ɊJ��)
    Call SubCtrlMovableCmd(ueppEnd)

    If strTitle = "" Then
        strTitle = "�V�X�e���G���[����"
    End If

    strMessage = DEFAULT_MESSAGE & strAddMessage

    vbmbrResult = MsgBox(strMessage, vbCritical + vbOKOnly, strTitle)

    '    '�d���̂Ő�Ƀ��b�Z�[�W
    '    Call SubCtrlMovableCmd(ueppEnd)


End Sub

'�G���[����
Sub SubHandleErrorAndFinishing(Optional ByVal strTitle As String = "")

    Select Case Err.Number
        Case 0
            Call Sub����I���\��(strTitle)             '����I�����܂Ƃ߂�����

        Case G_CTRL_ERROR_NUMBER_USER_NOTICE     '�G���[�ł͂Ȃ����ǒ���
            Call Sub����I���\��(strTitle, Err.Description)

        Case G_CTRL_ERROR_NUMBER_USER_CAUTION    '���[�U�[�G���[
            Call Sub���[�j���O�I���\��(strTitle, Err.Description)

        Case G_CTRL_ERROR_NUMBER_DEVELOPER       '�J���҃G���[
            Call Sub���[�j���O�I���\��(strTitle, Err.Description)

        Case Else                                '�V�X�e���G���[
            Call Sub�V�X�e���G���[����(strTitle, Err.Description)

    End Select

End Sub

'�n���ꂽ�t�H���_�p�X�̍Ōオ���łȂ�������t���ĕԂ�
Function FncChkPathEnd(ByVal strFolderPath As String) As String

    Const FOLDER_MARK   As String = "\"

    If Right(strFolderPath, 1) <> FOLDER_MARK Then
        strFolderPath = strFolderPath & FOLDER_MARK
    End If
    FncChkPathEnd = strFolderPath

End Function

'�n���ꂽ�t�@�C���p�X�̃t�H���_�p�X��Ԃ�
Function FncGetFolderNameFromFilePath(ByVal strFilePath As String) As String

    Dim myFSO       As FileSystemObject
    Dim strFolder   As String

    Set myFSO = New FileSystemObject

    strFolder = myFSO.GetParentFolderName(strFilePath)

    FncGetFolderNameFromFilePath = strFolder

    Set myFSO = Nothing

End Function

'�n���ꂽ�t�H���_�p�X�����݂��邩�`�F�b�N���Č��ʂ�Ԃ�
Function FncCheckFolderExist(ByVal strFolderPath As String) As Boolean

    Dim myFSO       As FileSystemObject
    Dim blnResult   As Boolean

    Set myFSO = New FileSystemObject

    blnResult = myFSO.FolderExists(strFolderPath)

    FncCheckFolderExist = blnResult

    Set myFSO = Nothing

End Function

'�n���ꂽ�t�@�C���p�X�����݂��邩�`�F�b�N���Č��ʂ�Ԃ�
Function FncCheckFileExist(ByVal strFilePath As String) As Boolean

    Dim myFSO       As FileSystemObject
    Dim blnResult   As Boolean

    Set myFSO = New FileSystemObject

    blnResult = myFSO.FileExists(strFilePath)

    FncCheckFileExist = blnResult

    Set myFSO = Nothing

End Function

'�t�@�C���R�s�[(�㏑���^)
Function FncCopySpecifFile(ByVal strCopyPath As String, ByVal strPastePath As String) As Boolean

    Dim myFSO       As FileSystemObject

    Set myFSO = New FileSystemObject

    myFSO.CopyFile strCopyPath, strPastePath, True

    FncCopySpecifFile = True                     '�V�X�e���G���[�����������_�ł����͒ʂ�Ȃ�

    Set myFSO = Nothing

End Function

'�t�@�C�����폜
Function FncDeleteFile(ByVal strFilePath As String) As Boolean

    Dim myFSO       As FileSystemObject

    Set myFSO = New FileSystemObject

    myFSO.DeleteFile strFilePath, True

    FncDeleteFile = True                         '�V�X�e���G���[�����������_�ł����͒ʂ�Ȃ�

    Set myFSO = Nothing

End Function

'�e�L�X�g�`���Ńt�@�C����ǂݍ��݁A�V�[�g��Ԃ�
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

'�w���Ń\�[�g��������(1�s�ڍ��ږ��A��������)
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

    '���Տ���
    With rngSortArea.Cells(1, 1)
        .Copy
        .PasteSpecial Paste:=xlPasteValues
    End With

    FncSortWithSpecifColumn = True
    Set WS = Nothing
    Set rngSortArea = Nothing

End Function

'�Z���͈͂���w��̕�����̃Z����Ԃ��i�Y���Ȃ���Nothing���Ԃ�j
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

'�w��̕�����̃Z����Ԃ��i�Y���Ȃ���Nothing���Ԃ�j
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

'�w��̕�����̃Z����Ԃ�Find�Łi�Y���Ȃ���Nothing���Ԃ�j
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

'2�����z����L�[�ʒu�Ń\�[�g���� (1������2��������)
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
        Err.Raise G_CTRL_ERROR_NUMBER_DEVELOPER, , "�z��ȊO�g�p�s��"
    ElseIf lngRank > 2 Then
        Err.Raise G_CTRL_ERROR_NUMBER_DEVELOPER, , "1�����z��2�����z��̂ݎg�p��"
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

'�w��s�{�b�g���̂���V�[�g����Ԃ�
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

'�w��s�{�b�g�͈͕̔ύX�i2�܂Ń\�[�g�w��j
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

    With myPivot                                 'External:=True�K�{
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

'�w��s�{�b�g�̂���Z���͈͂�Ԃ��i������Ȃ���΃G���[�I���j
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
        "�w��̃s�{�b�g�e�[�u��������܂���" & vbCrLf & _
        "�s�{�b�g�e�[�u���͍폜���ꂽ�����O���ύX���ꂽ�\��������܂�"
    End If

    Set myRange = myPivot.TableRange1

endRoutine:
    Set FncGetTableRangeOfPivot = myRange
    Set myRange = Nothing
    Set myPivot = Nothing
    Set WS = Nothing

    If Err.Number <> 0 Then
        Call SubHandleErrorAndFinishing("�s�{�b�g�ʒu�T��")
    End If

End Function

'�s�{�b�g�e�[�u���̃Z���͈͂�n���A�L����ԁi�l�������Ă���j���ǂ����Ԃ�(True���G���[���b�Z�[�W�t��False)
Function FncChekPivotRangeValid(ByVal rngPivot As Range, _
Optional ByVal blnRow As Boolean = True, _
Optional ByVal blnColumn As Boolean = True) As Boolean

    Const ERROR_NONE    As String = "�s�{�b�g�e�[�u���ɗL���f�[�^������܂���" & vbCrLf & _
    "�f�[�^��ǂݍ���ł���Ď��s���ĉ�����"
    Const ERROR_INVALID_ARGUMENT    As String = "Row��Column�̗���False�ɂ��邱�Ƃ͏o���܂���"

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
        'False���Ԃ�
        Err.Raise G_CTRL_ERROR_NUMBER_USER_CAUTION, , ERROR_NONE
    End If

End Function

'�z��̎��������擾����
Function FncGetArrayDimension(ByVal avntArray As Variant) As Long

    Dim d       As Long
    Dim tmp     As Long

    If Not IsArray(avntArray) Then Exit Function '0���Ԃ�

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

'�f�[�^���������ǂ����Ԃ��i�����͐��l�̂݁j
Function IsInteger(ByVal vntData As Variant) As Boolean

    Dim blnResult       As Boolean

    If IsNumeric(vntData) Then
        If Int(vntData) = vntData Then
            blnResult = True
        End If
    End If

    IsInteger = blnResult

End Function

'�n���ꂽ�f�[�^�������݂̂��ǂ�����Ԃ�(�����_�A�n�C�t���A�L���̓_��)
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

    If lngResult <> 0 Then                       '��������Ȃ����0�������Ă�
        blnResult = True
    End If

    FncCheckDataIsNumberOnly = blnResult

End Function

'�z��̒��ɃG���[�l�����邩
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

'������𕶎���z��ɕϊ����ĕԂ�
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
''�N���X�N���֐�
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


