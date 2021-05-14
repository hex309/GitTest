Attribute VB_Name = "mdl_Env"
Option Explicit
Option Private Module


'============================================================
'�@�t�@�C������E�t�@�C������
'============================================================

'�t�@�C���_�C�A���O���J���A�c�[���V�[�g�̃t�@�C���ݒ藓�Ƀ��[�U���I�������t�@�C���p�X������
Sub SubSetCSVFilePathByUserChoice(ByVal uexlSpecify As ueXLFileType)

    Const INITIAL_FILE    As String = "C:\"

    Dim myRange         As Range
    Dim strFolderPath   As String
    Dim strFilePath     As String

    '�t�@�C���_�C�A���O���J���̂Ŏ��{�m�F�͏o���Ȃ�
    Call SubCtrlMovableCmd(ueppStart)

    '�v���Z�b�g�Ƃ��ăt�@�C���ݒ藓�̃Z�����擾(Nothing�ԋp�͂Ȃ�)
    Set myRange = FncGetRangeOfFileSetting(uexlSpecify) '�������Ɏg���̂�Range�����
    If myRange.Value <> "" Then
        '�t�@�C���̃t�H���_���擾
        strFilePath = FncGetFolderNameFromFilePath(myRange.Value)
        strFolderPath = FncChkPathEnd(strFilePath)

        '�V�[�g�����󗓂̏ꍇ���̃t�@�C���̃t�H���_��
    Else
        strFolderPath = ThisWorkbook.Path
    End If


    '�����l�n���Ȃ���_�C�A���O�J��
    strFilePath = FncOpenDialogAndGetFile(strFolderPath)

    If strFilePath = "" Then                     '�󗓂̓L�����Z��
        Call Sub�L�����Z���\��
    Else

        Application.ScreenUpdating = True        '��������悤��
        myRange.Value = strFilePath

        Call Sub����I���\��(, "�t�@�C����ݒ肵�܂���")

    End If

    Set myRange = Nothing

End Sub

'�t�@�C���I������(�L�����Z����G���[�͋󔒂��Ԃ�)
Function FncOpenDialogAndGetFile(Optional ByVal strInitialPath As String = "") As String

    Const DEF_TITLE     As String = "CSV�t�@�C���̑I��"

    Dim strFilterName   As String
    Dim strFilterExt    As String
    Dim strGetResult    As String

    strFilterName = G_FILTERNAME_EXCEL
    strFilterExt = "*" & G_EXT_XLS

    '���s
    With Application.FileDialog(msoFileDialogFilePicker)

        '�����I���̉s��(���̃v���V�[�W���ł͕s����)
        .AllowMultiSelect = False

        '�t�B���^�̃N���A
        .Filters.Clear

        '�t�@�C���t�B���^�̒ǉ�
        .Filters.Add strFilterName, strFilterExt

        '�����\���t�@�C���̐ݒ�
        If strInitialPath <> "" Then
            .InitialFileName = strInitialPath
        End If

        '�_�C�A���O�^�C�g��
        .Title = DEF_TITLE

        '���ʎ擾
        If .Show = True Then
            strGetResult = .SelectedItems(1)
        Else
            strGetResult = vbNullString          '=""
        End If

    End With

    FncOpenDialogAndGetFile = strGetResult


End Function

'�t�@�C���ݒ藓�̃Z����Ԃ�(Nothing�ԋp�͂Ȃ�)
Function FncGetRangeOfFileSetting(ByVal uexlSpecify As ueXLFileType) As Range

    Const FOOTER_NAME       As String = "�t�@�C����"
    Const SAFE_ADDRESS_OA   As String = "D7"     '����_����
    Const SAFE_ADDRESS_SS   As String = "D12"    '���i�䒠�i�����j
    Const SAFE_ADDRESS_DD   As String = "D17"    '���i�䒠�i����j
    Const SAFE_ADDRESS_SU   As String = "D30"    '�d����

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

    '�t�@�C���ݒ藓������
    Set myRange = FncGetTagetRange(strTarget, G_SHEETNAME_TOOL, 0, 1)

    If myRange Is Nothing Then
        Set myRange = ThisWorkbook.Worksheets(G_SHEETNAME_TOOL).Range(strAddress)
    End If

    Set FncGetRangeOfFileSetting = myRange
    Set myRange = Nothing

End Function

'�c�[���V�[�g�̃t�@�C���ݒ���擾���ăt�@�C���L����True�Ȃ�t�@�C���p�X��Ԃ�
Function FncGetFileSetting(ByVal uexlSpecify As ueXLFileType) As String

    '�萔
    Const ERR_MESSAGE_N1    As String = "�t�@�C�����ݒ肳��Ă��܂���" & vbCrLf
    Const ERR_MESSAGE_N2    As String = "�t�@�C����ݒ肵�Ă���Ď��s���ĉ�����"
    Const ERR_MESSAGE_N3    As String = "���S�݌ɐ��͏o�܂���"
    Const ERR_MESSAGE_C1    As String = "�w��̃t�@�C�������݂��܂���" & vbCrLf
    Const ERR_MESSAGE_C2    As String = "�t�@�C���ݒ肩�t�H���_�����m�F����" & vbCrLf & _
    "�������t�@�C������ݒ肵�ĉ�����"

    '�ϐ�
    Dim myRange             As Range
    Dim strFilePath         As String
    Dim strErrorMessage     As String
    Dim lngErrNum           As Long
    Dim blnResult           As Boolean

    '�t�@�C���ݒ藓�̃Z�����擾(Nothing�ԋp�͂Ȃ�)
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
    '�����󂾂�����
    If myRange.Value = "" Then
        lngErrNum = G_CTRL_ERROR_NUMBER_USER_CAUTION
        strErrorMessage = TargetFile & ERR_MESSAGE_N1 & ERR_MESSAGE_N2
        GoTo endRoutine
    End If
    strFilePath = myRange.Value

    '�w��t�@�C���̗L���m�F
    blnResult = FncCheckFileExist(strFilePath)
    If blnResult = False Then
        lngErrNum = G_CTRL_ERROR_NUMBER_USER_CAUTION
        strErrorMessage = TargetFile & ERR_MESSAGE_C1 & ERR_MESSAGE_C2
        GoTo endRoutine
    End If

    FncGetFileSetting = strFilePath


endRoutine:
    Set myRange = Nothing
    If lngErrNum <> 0 Then                       'Err.Number�Ƃ͈Ⴄ�̂Œ���
        Err.Raise lngErrNum, , strErrorMessage
    End If

End Function

Public Sub GetFoldePath()
    On Error GoTo ErrHdl
    ThisWorkbook.Worksheets("���f�[�^�Ǎ�").Range(SAVE_FOLDER).Value = ShowFileDialog

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


