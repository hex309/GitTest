Attribute VB_Name = "RikuNaviOpe"
Option Explicit
Option Private Module

Public Function loginRikuNavi(ByRef argTgtSite As CorpSite, Optional msgFlg As Boolean = True) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    Dim i As Long
    Dim v As Object
    Dim ar As Variant
    Dim year As String
    
    tgtURL = ""
    
    With argTgtSite
        If Not .navigate(.baseURL & "rms/") Then GoTo loginErr
        
        .byName("kokyakuCd")(0).Value = .CorpID
        .byName("rmsUserCd")(0).Value = .userName
        .byName("pwd")(0).Value = .userPass
        .click .byName("doLogin")(0)
        
        If Not .byId("form1", , False) Is Nothing Then
            If InStr(.byId("form1").innerHTML, "alert_box_text") > 0 Then
                GoTo loginErr
            End If
        End If
        
        ar = Split(.baseURL, "/")
        year = ar(UBound(ar) - 1)
        
        For Each v In .byId("h_nav").getElementsByTagName("a")
            If InStr(v.innerText, year) > 0 Then
                .click v
                Exit For
            End If
        Next

        If Not .navigate(.baseURL & "rics/login") Then GoTo loginErr
        
        If Not .byId("main2", , False) Is Nothing Then
            If InStr(.byId("main2").innerText, "�G���[") > 0 Then
                opeLog.Add .byId("main2").innerText
                GoTo loginErr
            End If
        End If
        
        If msgFlg Then
            AlertBox.Label1.ForeColor = &HFF&
            AlertBox.Label1 = "�����F�؂���͂��Ă��������I"
        End If
        
        sendCaptAlert "���N�i�r�̃��O�C���F�ؑ҂��ł��B"
        
        On Error GoTo loginErr
        If Not .waitPageMoved("rics/top/") Then GoTo loginErr
        On Error GoTo 0
        
        If msgFlg Then AlertBox.Label1.ForeColor = &H80000012
        
    End With
    
        loginRikuNavi = True

Exit Function
loginErr:
    opeLog.Add "���N�i�r�Ƀ��O�C���ł��܂���ł����B" & vbCrLf & "�A�h���X/�A�J�E���g/�p�X���[�h/�ʐM�󋵂����m�F���������B"
    loginRikuNavi = False
        
End Function

Public Function searchRikuNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As Boolean
    Dim ankElmt As Object

    With argTgtSite
'##������ʂւ̑J�ڂƁA���������̓���
        If Not .navigate(.baseURL & "rics/search/condition/profile/doInit/") Then Exit Function

        If tgtDtType = dataType.personal Then
            If Not .navigate(.baseURL & "rics/search/condition/profile/") Then Exit Function
            .byName("tourokuDateFrom")(0).Value = Format(.lastUpdate, "yyyymmdd")
            .byName("tourokuDateKbn")(0).Value = 2
            .byName("tourokuDateTo")(0).Value = Format(Date, "yyyymmdd")
            .click .byName("doAdd")(0)
            
            '##���s�{�^����
            'JS�𒼐ڎ��s����ƃG���[�ɂȂ�
            '.execScript "RnWINKCommonSubmit(this, '/2020/rics/search/result/doSearch');"
            
            For Each ankElmt In .byId("main").getElementsByTagName("input")
                If ankElmt.className = "act_btn" Then
                    .click ankElmt
                    Exit For
                End If
            Next
        Else
            If Not .navigate(.baseURL & "rics/search/mySearch/") Then Exit Function
            
            .execScript "searchStudentFromAnchor('searchAllFlg');"
        End If

        If InStr(.byId("mainFull").innerText, "�����ɊY������w�������݂��܂���ł����B") Then
            opeLog.Add Format(.lastUpdate, "yyyy/mm/dd") & " �` " & Format(Date, "yyyy/mm/dd") & "�Ō������܂������Y������f�[�^�͂���܂���ł����B"
            Exit Function
        End If
        '###20200301 CMJ Nakabayashi�C��
        '###�_�E�����[�h�{�^�����A���X�g����I���ł͂Ȃ��A�����{�^���ɕ����ꂽ����
        '###�X�N���v�g�𒼐ڎ��s
        .objIE.Document.Script.setTimeout "javascript:setDownloadType(0);", 1000
       
        '.byId("action_select_summit").Value = "doSetDlTarget_summit_0"
        '.byId("execButton_summit").click
        
        '�G���[�`�F�b�N
        If Not .byId("main2", , False) Is Nothing Then
            If InStr(.byId("main2").innerText, "�G���[") > 0 Then
                opeLog.Add .byId("main2").innerText
                Exit Function
            End If
        End If
        
    End With
    
    searchRikuNaviDt = True

End Function

Public Function makeRikuNaviDt(ByRef argTgtSite As CorpSite, ByVal tgtDtType As Long) As String
    Dim table As Variant
    Dim fileName As String
    Dim ankElmt As Object
    Dim layOutName As String
    Dim a As String
    Dim i As Long
    Dim rdChkOk As Boolean
    Dim fmIptOk As Boolean
        
    
    With argTgtSite
    
        If tgtDtType = dataType.personal Then
            layOutName = .dlLayout
        Else
            layOutName = .dlLayoutEV
        End If
    
        For Each table In .byId("main").getElementsByTagName("table")
            For i = 0 To table.Cells.Length - 1
                If table.Cells(i).innerText = Trim(layOutName) Then
                    .click table.Cells(i - 1).Children(0)
                    rdChkOk = True
                End If
    
                If table.Cells(i).innerText = "�_�E�����[�h�t�@�C����" Then
                    fileName = .getFileNameNow("���N�i�r", tgtDtType)
                    table.Cells(i + 1).Children(0).Value = fileName
                    fmIptOk = True
                End If
            Next
        Next
        
        If Not (rdChkOk And fmIptOk) Then
            opeLog.Add "���N�i�r�̃��C�A�E�g���I���ł��Ȃ����A�t�@�C���������͂ł��܂���ł����B"
            Exit Function
        End If
        
        For i = 1 To 2
            For Each ankElmt In .byId("main").getElementsByTagName("input")
                If ankElmt.className = "act_btn" Then
                    .click ankElmt
                End If
            Next
        Next
        
    End With
    
    makeRikuNaviDt = fileName
    
End Function

Public Function waitRikuNaviCSV(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As Boolean
    Dim timeOut As Date
    Dim noAltFlg As Boolean
    Dim alLink As Object
    Dim i As Long
    
    timeOut = Now + SettingSh.DlTimeOut

    With argTgtSite
        Do While .getTableCell(argFileName, argFileName, -1).innerText = vbNullString
            Do While InStr(.byId("alert").innerText, "�_�E�����[�h") = 0
                DoEvents
                Application.Wait Now + TimeValue("00:00:03")
        
                If Now > timeOut Then
                    opeLog.Add "�_�E�����[�h�t�@�C�������̃A���[�g�����m�ł��܂���ł����B"
                    noAltFlg = True
                    
                    i = i + 1
                    If i > 1 Then
                        Exit Function
                    End If
                End If
                
                If cancelFlg Then
                    opeLog.Add "�L�����Z������܂����B"
                    Exit Function
                End If
            Loop
            
            If noAltFlg Then
                If Not .navigate(.baseURL & "rics/download/reservedList/showCsvList/") Then
                    Exit Function
                End If
            Else
                Set alLink = .byId("alert").getElementsByTagName("a")(0)
                .click alLink
            End If
        Loop
    End With
    
    waitRikuNaviCSV = True
        
End Function


Public Function dlRikuNaviCSV(ByRef argTgtSite As CorpSite, ByVal argFileName As String) As String
    Dim timeOut As Date
    Dim noAltFlg As Boolean
    Dim filePath As String
    Dim table As Variant
    Dim i As Long
    Dim tgtCell As Object
    Dim alrPass As Boolean
    
    timeOut = Now + SettingSh.DlTimeOut

    With argTgtSite
        Set tgtCell = .getTableCell(argFileName, argFileName, -1)
        tgtCell.Children(0).click
        
        If .byName("imageAuthKey")(0).Value <> vbNullString Then
            .execScript "checkValue();", 3
            
            If .byTag("h2")(0).innerText = "�G���[" Then
                If Not .navigate(.baseURL & "rics/download/reservedList/showCsvList/") Then
                    Exit Function
                End If
                Set tgtCell = .getTableCell(argFileName, argFileName, -1)
                tgtCell.Children(0).click
            Else
                alrPass = True
            End If
        End If
        
        If Not alrPass Then
            .byName("imageAuthKey")(0).Focus
            
            AlertBox.Label1.ForeColor = &HFF&
            AlertBox.Label1 = "�����F�؂���͂��ă{�^���������Ă��������I"
            
            sendCaptAlert "���N�i�r�̃_�E�����[�h�F�ؑ҂��ł��B"
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
                '�^�C���A�E�g�̃��b�Z�[�W�͌Ăяo����Ŕ���
                AlertBox.Label1.ForeColor = &H80000012
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "�L�����Z������܂����B"
                Exit Function
            End If
            
        Loop While filePath = vbNullString
        
    End With
    
    AlertBox.Label1.ForeColor = &H80000012
    dlRikuNaviCSV = filePath
End Function
