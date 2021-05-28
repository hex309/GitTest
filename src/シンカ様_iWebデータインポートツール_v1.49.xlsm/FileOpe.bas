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
        opeLog.Add srcCSVpath & vbCrLf & "��L�t�@�C��������܂���B"
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
        opeLog.Add srcCSVpath & vbCrLf & "��L�t�@�C��������܂���B"
        Exit Function
    End If
    
    titleLine = txtStrm.ReadLine
    
    KJ_FAM_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KJ_FAM_NAME")
    KJ_FST_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KJ_FST_NAME")
    KN_FAM_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KN_FAM_NAME")
    KN_FST_NAME_IDX = LabelSh.getCSVColumnIndex(titleLine, "KN_FST_NAME")
    
    targetIDXs = Array(KJ_FAM_NAME_IDX, KJ_FST_NAME_IDX, KN_FAM_NAME_IDX, KN_FST_NAME_IDX)

    Do
        newFilePath = srcCSVpath & "_�C����" & IIf(i = 0, "", "_" & i) & ".csv"
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
                dtLine = Replace(dtLine, csvCells(j), Replace(csvCells(j), "�@", vbNullString))
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
        opeLog.Add "�����������̓t���K�i�ɃX�y�[�X�̍��������m�������ߏ���"
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
    opeLog.Add "�t�@�C������ύX�ł��܂���ł����B" & vbCrLf & "�ύX���悤�Ƃ����t�@�C�����F" & retPath


End Function


Public Function getDlFilePath(ByVal fileName As String, Optional ByVal msgFlg As Boolean = True) As String
    Dim wsh As Object 'WshShell
    Dim fso As Object 'FileSystemObject
    Dim folderPath As String
    Dim filePath As String
    
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    folderPath = SettingSh.DlFolderPath
    
    If InStr(folderPath, "�_�E�����[�h") > 0 Then
        folderPath = Replace(folderPath, "�_�E�����[�h", "Downloads")
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
                opeLog.Add filePath & vbCrLf & vbCrLf & "�_�E�����[�h��Ɏw�肳�ꂽ�t�H���_�Ƀt�@�C��������܂���B(��L�p�X)" & vbCrLf _
                    & "InternetExplorer�̋K��̃_�E�����[�h�t�H���_�ƁA�{�c�[���Ń_�E�����[�h��Ɏw�肳�ꂽ�t�H���_����v���Ă��邩�m�F���Ă��������B" & vbCrLf & vbCrLf _
                    & "�܂��_�E�����[�h����t�@�C���T�C�Y���傫�����߂Ɏ��Ԃ��������Ă���ꍇ�́A�^�C���A�E�g���鎞�Ԃ��������Ă��������B" & vbCrLf _
                    & "���݂̃^�C���A�E�g�ݒ�ihh:mm:ss�j�F" & Format(SettingSh.DlTimeOut, "hh:mm:ss")
            End If
        End If
    Else
        If msgFlg Then
            opeLog.Add folderPath & vbCrLf & vbCrLf & "�{�c�[���Ń_�E�����[�h��Ɏw�肳�ꂽ�t�H���_�i��L�j������܂���B"
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
        opeLog.Add csvPath & vbCrLf & vbCrLf & "�t�@�C����������܂���B"
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
    
    opeLog.Add dPeople.allPeople.Count & "�l���̃f�[�^�����[�h���܂����B"

    'On Error GoTo 0
    
    Set getPeople = dPeople

normalFin:
    Set fso = Nothing
    Set txtStrm = Nothing
    
'    Exit Function
'Err:
'    MsgBox "CSV���󔒂��A�������͓ǂݎ��܂���B" & vbCrLf _
'        & csvPath & vbCrLf _
'        & "��L�t�@�C�����m�F���Ă��������B", vbExclamation
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
        '���l�A���t�ւ̏�����D�悷��Ɓu1,2,3�v���̃f�[�^���G�X�P�[�v�ł��Ȃ��̂ŕ�����ɑ΂��鏈����D��
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
        '�n�C�t���͓��t�ȊO�̉\��������̂ŏ��O
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
'//  �v���V�[�W�����@�FgetCurrentRegion
'//  �@�\�@�@�@�@�@�@�F�w�肳�ꂽRange����A�^�C�g���s��������CurrentRegion��Ԃ��B
'//  �����@�@�@�@�@�@�FbaseCellRange�F�N�_�ƂȂ�͈�
'//                    titleRowSize: �^�C�g���s�̍s����I�v�V������f�t�H���g��0
'//  �߂�l�@�@�@�@�@�F�^�C�g���s��������CurrentRegion�A�擾�ł��Ȃ����Nothing��Ԃ��B
'//  �쐬�ҁ@�@�@�@�@�FAkira Hashimoto
'//  �쐬���@�@�@�@�@�F2017/12/28
'//  ���l�@�@�@�@�@�@�F
'//  �X�V���F���e�@�@�F
'// ------------------------------------------------------------------------------------------------------------------------

Function getCurrentRegion(ByVal baseCellRange As Range, _
                          Optional ByVal titleRowSize As Long = 0, _
                          Optional ByVal needAlart As Boolean = True) As Range
    
    '�G���[���o���ꍇ�́A�uCurError�v�ɔ��
    On Error GoTo CurError
                         
    '�����\�J�n�Z��(baseCellRange)�A�����^�C�g���s��(titleRowSize)����ɁA�f�[�^���擾
    '�y�v�ύX�zResize��Offset����ւ��ɂ���
    With baseCellRange.CurrentRegion
        Set getCurrentRegion = .offset(titleRowSize, 0).Resize(.Rows.Count - titleRowSize, .Columns.Count)
    End With
    
    '�G���[���o��ꍇ�@=�@�l���Ȃ��ꍇ
    If getCurrentRegion Is Nothing Then
        err.Raise Number:=GET_CUR_REG_ERR, Description:="Can not get Current Region"
    End If
    
    '�l������ꍇ�A�G���[�𖳌��ɂ��ăv���V�[�W���𔲂���
    On Error GoTo 0
    Exit Function

'�G���[���o���ꍇ�̏���
CurError:
    'Range�I�u�W�F�N�g�ɁuNothing�v����
    Set getCurrentRegion = Nothing
    
    '����(needAlart)��True�̏ꍇ�A�A���[�g��\��
    If needAlart Then
        opeLog.Add baseCellRange.Parent.name & vbCrLf & "��L���[�N�V�[�g�̃f�[�^�͈͂��擾���o���܂���ł����B"
    End If
    
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  �v���V�[�W�����@�FsplitCSVLine
'//  �@�\�@�@�@�@�@�@�FCSV�̂P�s�����R���}�ŕ������ăR���N�V�����ŕԂ��܂��B��{�I�ɂ�Excel�œǂ񂾂̂Ɠ������ʂɂȂ�l���Ă��܂��B
'//  �@�@�@�@�@�@�@�@�@���P:�s�̐擪�A����уR���}�̌�̃_�u���N�H�[�e�[�V�����̂݃t�B�[���h�̃G�X�P�[�v�Ɖ��߂��܂��
'//  �@�@�@�@�@�@�@�@�@���Q:�t�B�[���h���G�X�P�C�v����Ă��Ȃ��Ƃ��A�_�u���N�H�[�e�[�V�����ŃG�X�P�C�v���Ă��Ȃ��_�u���N�H�[�e�[�V�����������Ɖ��߂��܂��
'//  �����@�@�@�@�@�@�FcsvLine�FCSV�̂P�s��
'//  �߂�l�@�@�@�@�@�FCSV�̃t�B�[���h��item�ł���Collection
'//  �쐬�ҁ@�@�@�@�@�FAkira Hashimoto
'//  �쐬���@�@�@�@�@�F2018/03/06
'//  ���l�@�@�@�@�@�@�F
'//  �X�V���F���e�@�@�F
'// ------------------------------------------------------------------------------------------------------------------------

'�R���N�V�����Ŏ擾
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
    
    '�ŏI�J�����o��
    splitCSVLine.Add colValue

End Function

'Sub test()
'    makeDiffFile "C:\Users\���c�\\Downloads\90017�X�o��_�}�C_�Z�~�i0425152457.txt", "C:\Users\���c�\\Downloads\90017�X�o��_�}�C_�Z�~�i0424085108.txt", #4/26/2019 10:00:00 AM#
'
'End Sub


'�����t�@�C�����o�͂��AlastUpdate���Â����t��lastUpdate�֏���������
Public Function makeDiffFile(ByVal newCSVPath As String, ByVal oldCSVpath As String, ByVal lastUpdate As Date) As String
    Dim outDirPath As String
    Dim olds As Collection
    Dim data As Collection
    Dim i As Long, j As Long
    
    Const TITLE_ROW As Long = 1
    
    '#CSV�@���@Collection��
    Set data = getData(newCSVPath, 0)
    Set olds = getData(oldCSVpath, 0)
    
    If data Is Nothing Then GoTo abnormalFin
        
    '#����f�[�^���폜�i�X�V���ꂽ�f�[�^���c���j
    '#OLD�������ꍇ�͔�΂�
    
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

    '#�ȉ��f�[�^����
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
    
    '#�L�����Z����̃^�C�g���A�L�����Z��������������ݒ�

    Const CANCEL_TITLE = "�L�����Z���t���O"
    Const CANCEL_TEXT = "1"
    Const UPDATE_TITLE = "�G���g���[����"
     
    '�f�[�^����肷�邽�߂̗�̑g�ݍ��킹�Bkey = ��^�C�g���Aitem = ��ԍ�
    Set uniqTitles = CreateObject("Scripting.Dictionary")
    
    uniqTitles.Add "�w���Ǘ�ID", 0
    uniqTitles.Add "�Z�~�i�[�ԍ�", 0
    uniqTitles.Add "�G���g���[����", 0

    '�^�C�g���s���Ȃ���΃G���[
    If TITLE_ROW < 1 Then
        MsgBox "�^�C�g���s���K�v�ł��B", vbExclamation
        GoTo abnormalFin
    End If
    
    Set targetCells = splitCSVLine(data(TITLE_ROW))
    
    '�ŐV�t�@�C���̃^�C�g������u�L�����Z���v�Ɓu�G���g���[�����v�A�̗�ԍ���T��
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
    
    '�݂���Ȃ���΃G���[
    If cancelColumn < 1 Then
        MsgBox "�ŐV�̃_�E�����[�h�f�[�^�Ɂw" & CANCEL_TITLE & "�x�񂪂���܂���B", vbExclamation
        GoTo abnormalFin
    End If
    
    If updateColumn < 1 Then
        opeLog.Add "�ŐV�̃_�E�����[�h�f�[�^�Ɂw" & UPDATE_TITLE & "�x�񂪂���܂���B"
        GoTo abnormalFin
    End If
    
    For Each targetCellValue In uniqTitles
        If uniqTitles(targetCellValue) = 0 Then
            opeLog.Add "�ŐV�̃_�E�����[�h�f�[�^�Ɂw" & targetCellValue & "�x�񂪂���܂���B"
            GoTo abnormalFin
        End If
    Next
      
    ' �t�@�C���o�͂̏����@�p�X�쐬�ƃe�L�X�g�J��
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With fso
    
    i = 0
    Do
        outPath = .BuildPath(.GetParentFolderName(newCSVPath), .GetBaseName(newCSVPath) & _
                                "_����" & IIf(i = 0, vbNullString, i) & ".csv")
        i = i + 1
    Loop While .FileExists(outPath)
    
    End With
    
    Set txtStrm = fso.CreateTextFile(outPath)


    '�^�C�g���s�������o��
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
    '�����f�[�^�������ꍇ�̓^�C�g���݂̂ňȉ��X�L�b�v
    '����ꍇ�́A
    For i = TITLE_ROW + 1 To data.Count
        Set targetCells = splitCSVLine(data(i))
        '�����Ȃ��ꍇ�̓}�b�`�t���O�@False
        MatchFlg = False
        
        For k = TITLE_ROW + 1 To oldCnt
            Set oldCells = splitCSVLine(olds(k))
            MatchFlg = True
            
            '�f�[�^����̑g�ݍ��킹����ł��Ⴆ�Δ�����
            For Each targetCellValue In uniqTitles
                tgtIdx = CLng(uniqTitles(targetCellValue))
                
                If oldCells(tgtIdx) <> targetCells(tgtIdx) Then
                    MatchFlg = False
                    Exit For
                End If
            Next
                       
            '�f�[�^�S�Ă���v�iMatchFlg��True�j�ŁuoldCells�v�����̃f�[�^�Ƃ��ă��[�v�𔲂���
            If MatchFlg Then
                Exit For
            End If
        Next
        
        Dim nologgedCancel As Boolean
               
'        '�G���g���[�������O��X�V�����Â��A���O��ˍ���ŃL�����Z����Ԃɕς�����ꍇ�́A�G���g���[������O��X�V���ɏ���������B
'        '�G���g���[�������O��X�V�����V�����u�\��v���O���������ꍇ�́A��L�L�����Z�����O�͏㏑������邪�A
'        '���̏ꍇ�u�\��v���O�����鎞�_�ł��ꂪ�ŐV�ł���̂ŁA�㏑����OK�B
'        If Not MatchFlg Then
'            nologgedCancel = targetCells(updateColumn) < lastUpdate And targetCells(cancelColumn) = CANCEL_TEXT
'        Else
'            nologgedCancel = targetCells(updateColumn) < lastUpdate And targetCells(cancelColumn) = CANCEL_TEXT And oldCells(cancelColumn) <> CANCEL_TEXT
'        End If
        
        '�i�}�C�i�r�j�L�����Z���̓^�C���X�^���v���X�V����Ȃ�
        
        If olds Is Nothing Then
            '�������Ƃ��Ă��Ȃ��ꍇ�́A�ǂꂪ�������킩��Ȃ��B
            '�̂ŁA�O��X�V��-1���ȍ~�̃��O��ΏۂƂ���B��������O�̃L�����Z���͑ΏۂƂ��Ȃ��B�i�Â��܂܂̓��t�ŏ��������j
            nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT And CDate(targetCells(updateColumn)) >= DateAdd("d", -1, lastUpdate)
        
        Else
            '�������Ƃ����ꍇ�A
            If MatchFlg Then
                '�O��ˍ���ŃL�����Z����Ԃɕς�����ꍇ��Unlogged�L�����Z���i�O��L�����Z���������ꍇ�͑Ή����p�j
                nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT And oldCells(cancelColumn) <> CANCEL_TEXT
            Else
                '�O��̃f�[�^���Ȃ��ꍇ�͂��̂܂�Unlogged�L�����Z���i�ŐV�̓��t�͂킩��Ȃ��j
                nologgedCancel = targetCells(cancelColumn) = CANCEL_TEXT
            End If
        End If
        
        '�f�[�^��CSV�ɖ߂�
        For j = 1 To targetCells.Count
            '�L�����Z���̏ꍇ�́A�^�C���X�^���v�����ĂɂȂ�Ȃ��̂ŁA�^�C���X�^���v��O��X�V��-1���ŏ㏑��
            '�f�[�^��M����������ł���Ă���̂ŁA�ő�1���x���ꍇ������B�܂荷���Ƃ��ďo�Ă����L�����Z���͏��Ȃ��Ƃ��O��X�V����1���O�ȍ~�ɕύX���ꂽ���̂ƌ����B
            '���̓�����iWeb���̗������ǂ��܂Œǂ����Ɍ����Ă���B
            If j = updateColumn And nologgedCancel Then
                If CDate(targetCells(updateColumn)) >= DateAdd("d", -1, lastUpdate) Then
                    outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & targetCells(j)
                Else
                    outLine = IIf(outLine = vbNullString, vbNullString, outLine & ",") & Format(DateAdd("d", -1, lastUpdate), "yyyy/m/d hh:mm:ss")
                End If
            '�L�����Z������ 1 �� 99�@�֕ύX����B�i���O���X�V����Ă��Ȃ��L�����Z���ƁA���N�i�r�̒ʏ�̃L�����Z�����������邽�߁j
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
        opeLog.Add csvPath & vbCrLf & "��L�t�@�C���͌�����܂���B"
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
'//  �v���V�[�W�����@�FgetFilePathByDialog
'//  �@�\�@�@�@�@�@�@�F�_�C�A���O���g�p���ă��[�U�[�ɒP��̃t�@�C����I�΂��A���̃p�X��Ԃ��B
'//  �����@�@�@�@�@�@�F�_�C�A���O�̃t�B���^�[�i�f�B�X�N���v�V�����A�g���q�j�ƃ^�C�g���B
'//  �@�@�@�@�@�@�@�@�@��������I�v�V�����
'//  �߂�l�@�@�@�@�@�F�t�H���_�̃p�X/vbNullString
'//  �쐬�ҁ@�@�@�@�@�FAkira Hashimoto
'//  �쐬���@�@�@�@�@�F2017/12/11
'//  ���l�@�@�@�@�@�@�F
'//  �X�V���F���e�@�@�F
'// ------------------------------------------------------------------------------------------------------------------------

Public Function getFilePathByDialog(Optional ByVal argExt As String = "*.*", _
                            Optional ByVal argDscr As String = "All files", _
                            Optional ByVal argTitle As String = "�t�@�C����I�����ĉ�����") As String
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
