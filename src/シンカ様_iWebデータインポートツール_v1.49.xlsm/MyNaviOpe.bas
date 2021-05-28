Attribute VB_Name = "MyNaviOpe"
Option Explicit
Option Private Module

Public Function loginMyNavi(ByRef argTgtSite As CorpSite) As Boolean
    With argTgtSite
        On Error GoTo loginErr
        If Not .navigate(.baseURL, , True) Then Exit Function
        
        .byId("mwCorpNo").Value = .CorpID
        .byId("empStaffPasswd").Value = .userPass
        .byId("tmpEmpStaffId").Value = .userName
        If Not .click(.byId("doLogin")) Then
            GoTo loginErr
        End If
        
        If .byId("errorsArea", , False) Is Nothing Then
            loginMyNavi = True
        Else
            opeLog.Add Trim(.byId("errorsArea").innerText)
            GoTo loginErr
        End If
        
        On Error GoTo 0
    End With
    

Exit Function

loginErr:
    opeLog.Add "�}�C�i�r�Ƀ��O�C���ł��܂���ł����B" & vbCrLf & "�A�h���X/�A�J�E���g/�p�X���[�h/�ʐM�󋵂����m�F���������B"
    loginMyNavi = False
        
End Function

Public Function moveSearchWindow(ByRef argTgtSite As CorpSite, Optional ByVal fromTop As Boolean = True) As Boolean
    With argTgtSite
        If fromTop Then
            'comMiwsTopLink�����������ȉ�ʂɑJ�ڂ����AsearchTop�������Ȃ��󋵂����܂ɔ�������Ƃ̂��ƁB
            '�����s�������A�ҋ@�����ėl�q���i2019/4/19�j
            Sleep 2000
            
            .click .byId("comMiwsTopLink")
            .click .byId("searchTop")
        Else
            .click .byId("searchTop")
        End If
    End With
    
    moveSearchWindow = True
End Function

Public Function searchMyNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As Boolean
    Dim ancElmt As Object
    Dim altMsg As String

    With argTgtSite
        If tgtDtType = dataType.personal Then
            If Not .click(.byId("searchDeliveryList")) Then Exit Function
            
            .byId("regdateBeforeYear").Value = Format(.lastUpdate, "yyyy")
            .byId("regdateBeforeMonth").Value = Format(.lastUpdate, "mm")
            .byId("regdateBeforeDay").Value = Format(.lastUpdate, "dd")
            .byId("regdateCond").Value = 2
            .byId("regdateAfterYear").Value = Format(Date, "yyyy")
            .byId("regdateAfterMonth").Value = Format(Date, "mm")
            .byId("regdateAfterDay").Value = Format(Date, "dd")
            
            For Each ancElmt In .byId("main").getElementsByClassName("actbtn")
                If ancElmt.id = "doRegist" Then
                    If Not .click(ancElmt) Then Exit Function
                    Exit For
                End If
            Next
            
             altMsg = Format(.lastUpdate, "yyyy/mm/dd") & " �` " & Format(Date, "yyyy/mm/dd") & "�Ō������܂������Y������f�[�^�͂���܂���ł����B"
            
        ElseIf tgtDtType = dataType.Seminar Then
            'Do Nothing �S������
            
             altMsg = "�����ꂩ�̓����ɗ\��̂���f�[�^�͂���܂���ł����B"
        Else
            Exit Function
        End If
        
        If Not .click(.byId("doSearch")) Then Exit Function
        
        If InStr(.byId("main").innerText, "�������ʂ�0���ł���") > 0 Then
            opeLog.Add altMsg
            Exit Function
        End If
        
        .byId("executeUpList").Value = "c_file_out_reserve_edit"
        If Not .click(.byId("doExecuteUp")) Then Exit Function
        
    End With
    
    searchMyNaviDt = True
    
End Function

Public Function moveSearchWindowOld(ByRef argTgtSite As CorpSite) As Boolean
    With argTgtSite

        '#�����N�����ǂ�p�^�[��
        .click .byId("comMiwsTopLink")
        .click .byId("searchTop")
        .click .byId("searchDeliveryList")
    End With
    
    moveSearchWindowOld = True

End Function

Public Function searchMyNaviPd(ByRef argTgtSite As CorpSite) As Boolean
    
    With argTgtSite
        .byId("regdateBeforeYear").Value = Format(.lastUpdate, "yyyy")
        .byId("regdateBeforeMonth").Value = Format(.lastUpdate, "mm")
        .byId("regdateBeforeDay").Value = Format(.lastUpdate, "dd")
        .byId("regdateCond").Value = 2
        .byId("regdateAfterYear").Value = Format(Date, "yyyy")
        .byId("regdateAfterMonth").Value = Format(Date, "mm")
        .byId("regdateAfterDay").Value = Format(Date, "dd")
        
        Dim ancElmt As Variant
    
        For Each ancElmt In .byId("main").getElementsByClassName("actbtn")
            If ancElmt.id = "doRegist" Then
                ancElmt.click
                Exit For
            End If
        Next
        
        IECheck .objIE
        
        .click .byId("doSearch")

        If InStr(.byId("main").innerText, "�������ʂ�0���ł���") > 0 Then
            opeLog.Add Format(.lastUpdate, "yyyy/mm/dd") & " �` " & Format(Date, "yyyy/mm/dd") & "�Ō������܂������Y������f�[�^�͂���܂���ł����B"
            Exit Function
        End If
        
        .docIE.getElementById("executeUpList").Value = "c_file_out_reserve_edit"
        .docIE.getElementById("doExecuteUp").click
        IECheck .objIE
        
    End With
    
    searchMyNaviPd = True

End Function

Public Function makeMyNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As String
    Dim table As Variant
    Dim dlLayout
    Dim fileName As String
    Dim tallyho As Boolean
    Dim i As Long
    
    With argTgtSite
        
        If tgtDtType = dataType.personal Then
            dlLayout = .dlLayout
        ElseIf tgtDtType = dataType.Seminar Then
            dlLayout = .dlLayoutEV
        Else
            Exit Function
        End If
    
        For Each table In .byTag("table")
            If table.Cells(0).innerText = "�I��" Then
                For i = 0 To table.Cells.Length - 1
                    If table.Cells(i).innerText = dlLayout Then
                        table.Cells(i - 1).Children(0).click
                        tallyho = True
                        Exit For
                    End If
                Next
                Exit For
            ElseIf InStr(table.innerText, "�G���[") > 0 Then
                opeLog.Add Replace(.byId("main").innerText, vbCrLf, vbNullString)
                Exit Function
            End If
        Next
        
        If Not tallyho Then
            opeLog.Add "�w" & dlLayout & "�x�Ƃ������}�C�i�r���C�A�E�g���͂���܂���B"
            Exit Function
        End If
        
        fileName = .getFileNameNow("�}�C�i�r", tgtDtType)
        
        .docIE.getElementById("name").Value = fileName
        .docIE.getElementById("downloadFormat[v_10]").click
        
        If Not .click(.byId("doConfirm")) Then Exit Function
        If Not .click(.byId("doRegist")) Then Exit Function
        
        .quitIE
        
        If Not .click(.byId("iOOutExe")) Then Exit Function
        
    End With
    
    makeMyNaviDt = fileName
    
End Function

Public Function chkDateCreated(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As Boolean
    Dim timeOut As Date

    timeOut = Now + SettingSh.DlTimeOut
    
    With argTgtSite
        Do While InStr(.getTableCell("�I��", argFileName, 3).innerText, "�쐬��") > 0
            .Refresh
            
            If Now > timeOut Then
                opeLog.Add "�f�[�^�o�͒��ɂɃ^�C���A�E�g���܂����B"
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "�L�����Z������܂����B"
                Exit Function
            End If
            
            Sleep 5000
            DoEvents
        Loop
    End With
    
    chkDateCreated = True
End Function

Public Function dlMyNaviCSV(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As String
    Dim timeOut As Date
    Dim filePath As String
    Dim alrPass As Boolean
    
    timeOut = Now + SettingSh.DlTimeOut
    
    With argTgtSite
        .getTableCell("�I��", argFileName, -1).Children(0).click
        
        If .byId("authId").Value <> vbNullString Then
            .click .byId("doOutput"), , 3
            
            If .byClass("midashi_1")(0).innerText = "�G���[" Then
                If Not .click(.byId("iOOutExe")) Then Exit Function
            Else
                alrPass = True
            End If
        End If
        
        If Not alrPass Then
            .byId("authId").Focus
            
            AlertBox.Label1.ForeColor = &HFF&
            AlertBox.Label1 = "�����F�؂���͂��ă{�^���������Ă��������I"
            
            sendCaptAlert "�}�C�i�r�̃_�E�����[�h�F�ؑ҂��ł��B"
        End If
        
        Do
            DoEvents
            
            On Error Resume Next
            pushSaveButton .objIE
            On Error GoTo 0
        
            'Application.Wait [now() + "00:00:00.5"]
            filePath = getDlFilePath(argFileName & ".txt", False)
            
            If Now > timeOut Then
                getDlFilePath (argFileName & ".txt")
'                '�^�C���A�E�g�̃��b�Z�[�W�͌Ăяo����Ŕ���
'                AlertBox.Label1.ForeColor = &H80000012
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "�L�����Z������܂����B"
                Exit Function
            End If
            
        Loop While filePath = vbNullString
        
    End With
    
'    AlertBox.Label1.ForeColor = &H80000012
    dlMyNaviCSV = filePath
        
End Function

Public Function getDiffFile(ByVal mynavSmDataFilePath As String, ByVal tgtCorpName As String, ByVal lastUpdate As Date) As String

    ' �O��̃Z�~�i�[�t�@�C���p�X�̍s�ԍ�
    Dim oldPathRow As Long
    
    oldPathRow = SettingSh.OldMyNaviRowIndex(tgtCorpName)
    
    If oldPathRow = 0 Then
        opeLog.Add SettingSh.name & "�V�[�g�ɑΏۊ�Ɩ��w" & tgtCorpName & "�x������܂���B"
        Exit Function
    End If

    ' �O��̃Z�~�i�[�t�@�C���p�X���݃`�F�b�N�p
    Dim CheckPath As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    CheckPath = SettingSh.Cells(oldPathRow, 2).Value
            
    '�p�X���󔒂��A�L�ڂ���Ă��Ă��t�@�C��������Ƃ��́A�����t�@�C������
    '�󔒂̏ꍇ�͍����t�@�C���������ɓ��t����lastUpdate�ɍX�V�����B
    If CheckPath = vbNullString Or fso.FileExists(CheckPath) Then
        getDiffFile = makeDiffFile(mynavSmDataFilePath, CheckPath, lastUpdate)
        
        If getDiffFile = vbNullString Then Exit Function
        
        SettingSh.Cells(oldPathRow, 3).Value = mynavSmDataFilePath
    Else
    '�p�X�̋L�ڂ�����A�t�@�C���������ꍇ�̓G���[
        opeLog.Add SettingSh.name & "��[B" & oldPathRow & "]�ɋL�ڂ��ꂽ�p�X�Ƀt�@�C�������݂��܂���"
    End If
    
End Function

Public Function getMynaviCancelDate(ByRef argTgtSite As CorpSite, _
                                    ByVal myNaviId As String, _
                                    ByVal semID As String, _
                                    Optional ByVal tabIndex As Long) As Date

    With argTgtSite
        .byId("searchCode", tabIndex).Value = myNaviId
        .click .byId("doSearchMenu", tabIndex)
        .click .byId("entryLink", tabIndex)
        
        IECheck .objIE
        
        Dim linkTag As Object
        
        For Each linkTag In .byId("linktab").getElementsByTagName("a")
            If InStr(linkTag.innerText, "������E�ʐ�") > 0 Then
                .click linkTag
                Exit For
            End If
        Next
        
        Dim table As Variant
        Dim update As Date
        Dim cancelDate As Date
        Dim tempDate As Date
        Dim i As Long
        
        For Each table In .byTag("table")
            If InStr(table.innerText, "������E�ʐڗ\��󋵈ꗗ") > 0 Then
                For i = 0 To table.Cells.Length - 1
                    If Trim(table.Cells(i).innerText) = semID Then
                        tempDate = CDatePlus(Trim(table.Cells(i + 5).innerText))
                        If tempDate >= update Then
                            update = tempDate
                            cancelDate = CDatePlus(Trim(table.Cells(i - 1).innerText))
                        End If
                    End If
                Next
                Exit For
            End If
        Next
        
    End With
    
    getMynaviCancelDate = cancelDate
    
End Function

Public Function CDatePlus(ByVal argStr As String) As Date
    Dim buf As String
    
    buf = Trim(argStr)
    buf = Replace(buf, vbCr, vbNullString)
    buf = Replace(buf, vbLf, vbNullString)
    
    If buf = vbNullString Then
        CDatePlus = 0
    Else
        CDatePlus = CDate(buf)
    End If
    
End Function
