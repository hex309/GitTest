Attribute VB_Name = "iwebOpe"
Option Explicit
Option Private Module

Public Function loginiWeb(argIWeb As CorpSite) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    
    tgtURL = "login/check"
    
    With argIWeb
        pstData = "id=" & .userName & "&pass=" & .userPass
        If Not .post(.baseURL & tgtURL, , pstData, True) Then GoTo loginErr '���O�C��
        If InStr(.objIE.LocationURL, "top/index") = 0 Then GoTo loginErr
        
        
'        If InStr(.byId("front").innerText, "tabcd = ""all"";") = 0 Then
'            .click .byId("taball")
'            Do While InStr(.byId("front").innerText, "tabcd = ""all"";") = 0
'                Sleep 1000
'            Loop
'        End If
    End With
    
    loginiWeb = True
    
Exit Function
loginErr:
    opeLog.Add "iWeb�Ƀ��O�C���ł��܂���ł����B" & vbCrLf & "�A�h���X/�A�J�E���g/�p�X���[�h/�ʐM�󋵂����m�F���������B"
    loginiWeb = False

End Function

Public Function selectTab(argIWeb As CorpSite, ByVal jobTabName As String) As Boolean
    Dim hitFlg As Boolean
    Dim aTag As Object

    With argIWeb
        If .byId("crnt").innerText = jobTabName Then
            selectTab = True
            Exit Function
        End If
        
        For Each aTag In .byId("tabArea").getElementsByTagName("a")
            If aTag.innerText = jobTabName Then
                .click aTag
                hitFlg = True
                Exit For
            End If
        Next
        
        If Not hitFlg Then Exit Function
    
        Do While .byId("loadArea").innerText = "now loading..."
            Sleep 100
        Loop
    End With
    
    selectTab = True
         
End Function

Public Function DLiWebData(argIWeb As CorpSite, ByVal dataTypeCode As Long) As String
    Dim objIE As Object 'InternetExplorer
    Dim tgtURL As String
    Dim pstData As String
    
    opeLog.Add "iWeb�Ƀ��O�C���J�n.. �A�J�E���g�F" & argIWeb.userName
    If Not loginiWeb(argIWeb) Then Exit Function
    opeLog.Add "iWeb�Ƀ��O�C�������B�l���_�E�����[�h�y�[�W�ֈړ��J�n.."
    
    With argIWeb
    
        tgtURL = "download/confirm/"
        pstData = "layflg=cd&dlcd=" & dataTypeCode & "&ptflg=all"
        
        If Not .post(.baseURL & tgtURL, , pstData) Then Exit Function
        opeLog.Add "�ړ������BCSV�t�@�C�����m�F��..."
        
        Dim Text As Variant
        Dim fileName As String
    
        For Each Text In Split(.byId("selectTable").innerText, vbCrLf)
            If InStr(Text, "csv") > 0 Then
                fileName = Trim(Text)
                Exit For
            End If
        Next
    
        If fileName = vbNullString Then
            opeLog.Add "�_�E�����[�h����CSV�̃t�@�C�������m�F�ł��܂���ł����B"
            Exit Function
        End If
        
        opeLog.Add "CSV�t�@�C�����m�F�B�_�E�����[�h�J�n..."
        .click .byId("btnDownload"), SettingSh.DlTimeOut, 3
    
        Dim timeOut As Date
        
        timeOut = Now + SettingSh.DlTimeOut
    
        Do
            On Error Resume Next
            pushSaveButton .objIE
            On Error GoTo 0
        
            Application.Wait Now + TimeValue("00:00:01")
            DLiWebData = getDlFilePath(fileName, False)
            
            If Now > timeOut Then
                getDlFilePath (fileName)
                '�^�C���A�E�g�̃��b�Z�[�W�͌Ăяo����Ŕ���
                Exit Function
            End If
            
            If cancelFlg Then
                opeLog.Add "�L�����Z������܂����B"
                Exit Function
            End If
        
        Loop While DLiWebData = vbNullString
    End With
    
    opeLog.Add "�_�E�����[�h�����B"
    
End Function

Private Function loadNayoseTb(ByVal argTable As Object) As Object
    Dim cell As Variant

    If argTable.className = "eventTableWidth700" Then
    
        Set loadNayoseTb = CreateObject("Scripting.Dictionary")
            
        For Each cell In argTable.Cells
        
            If cell.tagName = "TH" Then
                loadNayoseTb.Add cell.innerText, cell.nextElementSibling.innerText
                
            ElseIf cell.cellIndex = 0 Then
                loadNayoseTb.Add "�I��", cell
            ElseIf InStr(cell.innerText, "��w") > 0 Or InStr(cell.innerText, "�w�Z") Then
                loadNayoseTb.Add "�w�Z", cell.innerText
            ElseIf InStr(cell.innerText, "HMI") > 0 And Len(cell.innerText) = 10 Then
                loadNayoseTb.Add "ID", cell.innerText
            End If
        Next
    End If

End Function

Private Function isAllOvreWrite(ByVal nyTable As Object, ByVal navTable As Object) As Boolean
    Dim cell As Variant
    
    isAllOvreWrite = True
    
    For Each cell In nyTable
        If cell <> "ID" And nyTable(cell) <> navTable(cell) Then
            opeLog.Add IIf(Trim(nyTable(cell)) = vbNullString, "��", "(i-Web)�F" & nyTable(cell)) & _
                        "�� (�i�r)�F" & IIf(Trim(navTable(cell)) = vbNullString, "��", navTable(cell))
            If nyTable(cell) = " " Then
                isAllOvreWrite = isAllOvreWrite And True
            ElseIf cell = "�w�Z" And InStr(nyTable(cell), "���̑�") > 0 Then
                isAllOvreWrite = isAllOvreWrite And True
            Else
                isAllOvreWrite = False
            End If
            
        End If
    Next
End Function

Private Function checkOption(ByRef argOptionCell As Object, ByVal argID As String) As Boolean
    Dim opt As Variant

    For Each opt In argOptionCell.Children
        If opt.id = argID Then
        
            If opt.Checked = False Then
                opt.click
                Exit For
            End If
        End If
    Next

End Function

Private Sub execNayose(argIWeb As CorpSite)
    Dim nyTable As Object 'Dictionary '
    Dim imptTable As Object 'Dictionary '
    Dim iwebTables As Collection
    Dim tb As Object
    
    Set imptTable = CreateObject("Scripting.Dictionary")
    Set iwebTables = New Collection
    
'##���񂹉�ʂ���A�e�[�u����ǂݍ���
    For Each tb In argIWeb.byTag("table")
        Set nyTable = loadNayoseTb(tb)
        
        If Not nyTable Is Nothing Then
            If InStr(tb.innerText, "�X�V(�}�̂̂�)") > 0 Then
                Set imptTable = nyTable
            Else
                iwebTables.Add nyTable
            End If
            
            Set nyTable = Nothing
        End If
    Next
    
    Dim chk As String
    
    chk = "chkbx_0"
    Set nyTable = iwebTables(1)
    
    opeLog.Add "���w" & imptTable(" ����") & "�x�𖼊񂹂��܂��B"
    
'##�e�[�u���̗L���m�F
    If iwebTables.Count = 0 Then
        opeLog.Add "�����₪����܂���B"
        Exit Sub
    ElseIf iwebTables.Count > 1 Then
        opeLog.Add "�����̖��񂹌�₪����܂��I"
'        chk = "chkbx_" & iwebTables.Count - 1
'        Set nyTable = iwebTables(iwebTables.Count)
        Exit Sub
    End If

'##�㏑���Ώۂƕ��@�̑I��
    Dim kubun As String
    
    If isAllOvreWrite(nyTable, imptTable) Then
        kubun = "nkbn2"
        opeLog.Add "��L���AID : " & nyTable("ID") & " ��S�čX�V���܂��B"
    Else
        kubun = "nkbn4"
        opeLog.Add "��L���AID : " & nyTable("ID") & " �̔}�̂̂ݍX�V���܂��B"
    End If
    
    checkOption imptTable("�I��"), kubun
    checkOption nyTable("�I��"), chk '"chkbx_0"

'##�㏑�����{
    Dim btnVal As Variant
    
    For Each btnVal In Array("�m�@�F", "���@�s")
        For Each tb In argIWeb.byClass("s colored")
            If tb.Value = btnVal Then
                argIWeb.click tb
            End If
        Next
    Next
    
    Set nyTable = Nothing
    Set imptTable = Nothing
    Set iwebTables = Nothing
    
End Sub

Private Function sendMail(argIWeb As CorpSite) As Boolean
    Dim table As Variant
    Dim cntCells As Long
    Dim timeOut As Date
    Dim bufStr As String
    
    timeOut = SettingSh.DlTimeOut
    
    opeLog.Add "���[���̗\����J�n���܂��B"
    
    With argIWeb
        If Not .submit(.byName("form1")(0), timeOut) Then GoTo err
        If Not .submit(.byName("formMake")(0), timeOut) Then GoTo err
        If Not .click(.byId("timeset"), timeOut) Then GoTo err
        
        For Each table In .byTag("table")
            If table.Cells(0).innerText = "�\��\����" Then
                cntCells = table.Cells.Length
            
                If cntCells >= 13 Then
                    If Not .click(table.Cells(12).Children(0)) Then GoTo err
                Else
                    If Not .click(table.Cells(cntCells - 1).Children(0)) Then GoTo err
                End If
                
                Exit For
            End If
        Next
        
        '�t�H�[���Ɏ��������f�����܂őҋ@
        
        timeOut = Now + timeOut
        
        Do
            On Error Resume Next
            bufStr = .byId("timeno", , False).Value
            On Error GoTo 0
            
            If bufStr <> vbNullString Then
                Exit Do
            ElseIf Now > timeOut Then
                opeLog.Add "���[�����M�����ݒ�Ń^�C���A�E�g���܂���"
                GoTo err
            End If
            
            DoEvents
        Loop
        
        .byId("mailsflag-3").click
        
        If Not .click(.byId("btnConfirm"), timeOut) Then GoTo err
        If Not .click(.byId("btnComplete"), timeOut) Then GoTo err
    End With
    
    opeLog.Add "���[���̗\�񂪊������܂���"
    sendMail = True
Exit Function

err:
    opeLog.Add "���[���̗\��Ɏ��s���܂���"
    
End Function

Public Function ULPersonalData(argIWeb As CorpSite, ByVal filePath As String, ByVal iWebNo As Long) As Boolean
    Dim csvFiles As Collection
    Dim csvFilePath As Variant
    
    If Not loginiWeb(argIWeb) Then Exit Function
    
    filePath = verifyCSV(filePath)
    Set csvFiles = csvDivider(filePath, 500)
    
    If csvFiles.Count = 0 Then
        ULPersonalData = False
        Exit Function
    Else
        ULPersonalData = True
    End If
    
    For Each csvFilePath In csvFiles
        opeLog.Add "�����A�b�v���[�h�J�n�F" & csvFilePath
        ULPersonalData = ULPersonalData And ULPersonalOneFile(argIWeb, csvFilePath, iWebNo)
    Next
    
End Function

Private Function ULPersonalOneFile(argIWeb As CorpSite, ByVal filePath As String, ByVal iWebNo As Long) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    Dim i As Long
    Dim timeOut As Date
    Dim msg As String
    
    timeOut = SettingSh.DlTimeOut
    
    With argIWeb
        tgtURL = "wdi/index"
        pstData = ""
        
        If Not .post(.baseURL & tgtURL, , pstData) Then
            opeLog.Add "i-Web �ꊇ�C���|�[�g��ʂւ̑J�ڒ��ɖ�肪�������܂����B"
            Exit Function
        End If
    
        .byId("wdilayoutno").Value = iWebNo
        .byId("headerflg").click
        
        If Not setDialogByVBS Then Exit Function
        
        '�����̑҂����Ԃ��Z���Ǝ��̃t�@�C�������_�C�A���O�ɃC���v�b�g���铮��ŃG���[���ł�B
        DoEvents
        Sleep 500
    
        setFileName filePath
        
        Do While .byName("wdifile")(0).Value = vbNullString
             DoEvents
             If Now > timeOut Then Exit Do
        Loop
        
        Dim tgtElmt As Object
        Dim naviOk As Boolean
        
        '�擪�f�[�^�m�F�y�[�W�ւ̑J��
        For Each tgtElmt In .byName("form_up")
            If InStr(tgtElmt.Action, "importfirstconfirm") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "�擪�f�[�^�̂��m�F�����肢���܂��B") > 0 Then
                        naviOk = True
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " ����荞�߂܂���B" & vbCrLf & "�擪�f�[�^�m�F�y�[�W�ɑJ�ڂł��܂���B"
            Exit Function
        Else
            naviOk = False
        End If

        '�捞���e�m�F�y�[�W�ւ̑J��
        For Each tgtElmt In .byName("form1")
            If InStr(tgtElmt.Action, "importsecondconfirm") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "�������e�̂��m�F�����肢���܂�") > 0 Then
                        naviOk = True
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " ����荞�߂܂���B" & vbCrLf & "�������e�̊m�F�y�[�W�ɑJ�ڂł��܂���"
            Exit Function
        Else
            naviOk = False
        End If
       
        '���̓`�F�b�N���ʂ̃y�[�W�ւ̑J��
        For Each tgtElmt In .byName("form1")
            If InStr(tgtElmt.Action, "filecheck") > 0 Then
                If .submit(tgtElmt, timeOut) Then
                    If InStr(.byClass("navigation_top")(0).innerText, "�f�[�^�`�F�b�N���e�̂��m�F�����肢���܂�") > 0 Then
                        naviOk = True
                    ElseIf InStr(.byClass("navigation_top")(0).innerText, "�G���[�f�[�^�����݂��܂�") > 0 Then
                        opeLog.Add "����荞�݃f�[�^�ɃG���[�f�[�^�����݂��܂�� �蓮�Ŏ�荞�ނ��A�G���[�f�[�^���������Ă��������B"
                    End If
                End If
                
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " ����荞�߂܂���B" & vbCrLf & "���̓`�F�b�N���ʂ̃y�[�W�ɑJ�ڂł��܂���B"
            Exit Function
        Else
            naviOk = False
        End If
        
        '�o�^�����y�[�W�ւ̑J��
        For Each tgtElmt In .byName("form2")
            If InStr(tgtElmt.Action, "complete") > 0 Then
                If .submit(tgtElmt, timeOut) Then naviOk = True
                Exit For
            End If
        Next
        
        If Not naviOk Then
            opeLog.Add filePath & " ����荞�߂܂���B" & vbCrLf & "�o�^�����̃y�[�W�ɑJ�ڂł��܂���B"
            Exit Function
        Else
            naviOk = False
        End If
        
        Dim topMsg As String
        Dim formName As String
        
        '����/�����`�F�b�N
        Do While True
            Set tgtElmt = .byId("contents").getElementsByTagName("table")
            
            If tgtElmt.Length = 0 Then
                topMsg = .byId("contents").innerText
            Else
                topMsg = tgtElmt(0).innerText
            End If
            
            Set tgtElmt = Nothing
            
            If InStr(topMsg, "�������������Ă��܂�") > 0 Then
                msg = " �V�K�o�^�F" & .getTableCell("�V�K�o�^����", "�V�K�o�^����", 1).innerText & "�� /" _
                     & " �X�V�����F" & .getTableCell("�X�V����", "�X�V����", 1).innerText & "�� /" _
                     & " �����f�[�^�����F" & .getTableCell("�����f�[�^����", "�����f�[�^����", 1).innerText & "��"
            
                Exit Do
            ElseIf InStr(topMsg, "���[�����M����o�^�������܂���") > 0 Then
                Exit Do
            
            ElseIf InStr(topMsg, "���[�����M") > 0 Then
                msg = " �V�K�o�^�F" & .getTableCell("�V�K�o�^����", "�V�K�o�^����", 1).innerText & "�� /" _
                     & " �X�V�����F" & .getTableCell("�X�V����", "�X�V����", 1).innerText & "�� /" _
                     & " �����f�[�^�����F" & .getTableCell("�����f�[�^����", "�����f�[�^����", 1).innerText & "��"
            
                If .MailFlg Then
                    sendMail argIWeb
                Else
                    Exit Do
                End If
                
            ElseIf InStr(topMsg, "�폜") > 0 Then
                opeLog.Add "��荞�݃f�[�^�̃G���[�ɂ��A�C���|�[�g���������܂���ł����B"
                Exit Function
            
            ElseIf InStr(topMsg, "����") > 0 Then
                If InStr(topMsg, "�X�V�������܂����B") > 0 Then
                    formName = "form"
                Else
                    formName = "form1"
                End If
    
                .submit .byName(formName)(0), timeOut
                
                execNayose argIWeb
            Else
                opeLog.Add "���炩�̌����ŃC���|�[�g���������܂���ł����B" & vbCrLf _
                         & "i-Web��ʂ̃g�b�v���b�Z�[�W:" & IIf(topMsg = vbNullString, "��", topMsg)
                Exit Function
            End If
        Loop
    End With
        
    opeLog.Add msg
    mailInfo.Add msg
    
    ULPersonalOneFile = True
    
End Function

Public Function ULSeminarData(argIWeb As CorpSite, _
                                ByVal seminarJob As String, _
                                ByVal seminarName As String, _
                                ByVal uploadText As String, _
                                ByVal cancelFlg As Boolean) As Boolean
    Dim tgtURL As String
    Dim pstData As String
    Dim classEl As Variant
    Dim param As Variant
    Dim msgText As String
    
    If uploadText = vbNullString Then
        'opeLog.Add "��" & seminarName & "��" & IIf(cancelFlg, "�L�����Z��", "�o�^") & "����ID�͂���܂���B"
        ULSeminarData = True
        Exit Function
    End If
    
    If Not loginiWeb(argIWeb) Then Exit Function
    If Not selectTab(argIWeb, seminarJob) Then
        opeLog.Add "���y���s�zi-Web�g�b�v��ʂɁw" & seminarJob & "�x�^�u��������܂���B" & vbCrLf & "�ȉ���" & IIf(cancelFlg, "ID���L�����Z��", "ID + [TAB] +�C�x���gNo �͓o�^") & "�ł��Ă���܂���I�I"
        opeLog.Add uploadText
        Exit Function
    End If
    
    With argIWeb
        For Each classEl In .byClass("clearfix")
            If classEl.innerText = seminarName Then
                param = Split(Split(Mid(classEl.outerHTML, InStr(classEl.outerHTML, "menu/event/") + 11), "'")(0), "/")
                pstData = param(0) & "=" & param(1) & "&" & param(2) & "=" & param(3)
            End If
        Next
        
        If pstData = vbNullString Then
            opeLog.Add "���y���s�zi-Web�g�b�v��ʂ́w" & seminarJob & "�x�^�u���Ɂw" & seminarName & "�x��������܂���B" & vbCrLf & "�ȉ���" & IIf(cancelFlg, "ID���L�����Z��", "ID + [TAB] +�C�x���gNo �͓o�^") & "�ł��Ă���܂���I�I"
            opeLog.Add uploadText
            Exit Function
        End If
    
        tgtURL = "reserveimport/make"
        .post .baseURL & tgtURL, , pstData
        
        If cancelFlg Then
            .byId("eventdayno").Value = 99999
            .byId("matchingkey-1").click
            msgText = "��" & seminarJob & "/" & seminarName & "����ȉ���ID���L�����Z�����܂����B" & vbCrLf
        Else
            .byId("matchingkey-2").click
            msgText = "��" & seminarJob & "/" & seminarName & "�Ɉȉ���ID��\�񂵂܂����B" & vbCrLf
        End If
        
        .byName("ucode").ucode.Value = uploadText
        .execScript "onSubmit();"
               
        If Trim(.getTableCell("�G���[����", "�G���[����", 2).innerText) <> "0��" Then
            opeLog.Add "���y���s�z" & seminarJob & "/" & seminarName & "�ւ̃A�b�v���[�h�Ɏ��s���܂����B"
            opeLog.Add uploadText
            Exit Function
        End If
        
        .execScript "onSubmit();"

        '���O�o�͐��`
        Dim uploadArray As Variant
        Dim uploadDict As Object
        Dim v As Variant
        Dim pid As String
        Dim did As String
        Dim i As Long
        
        If Not cancelFlg Then
            Set uploadDict = CreateObject("Scripting.Dictionary")
            
            uploadArray = Split(uploadText, vbCrLf)
            uploadText = vbNullString
            
            For i = LBound(uploadArray) To UBound(uploadArray)
                pid = Split(uploadArray(i), vbTab)(0)
                did = Split(uploadArray(i), vbTab)(1)
                
                If Not uploadDict.Exists(did) Then
                    uploadDict.Add did, pid
                Else
                    uploadDict(did) = uploadDict(did) & vbCrLf & pid
                End If
            Next
            
            For Each v In uploadDict
                uploadText = uploadText & IIf(uploadText = vbNullString, vbNullString, vbCrLf) & "�C�x���gNo:" & v & vbCrLf & uploadDict(v)
            Next
        End If
        
        opeLog.Add msgText & uploadText
        'mailInfo.Add msgText & uploadText
    End With
    
    ULSeminarData = True
    
End Function

Function ULAllSeminarData(ByRef iweb As CorpSite, ByRef myNavi As CorpSite, _
                          ByVal myNavSeminor As people, ByVal rkNavSeminar As people, ByVal iwebPeople As people) As Boolean
    Dim pid As Variant
    Dim semiList As SeminarList
    Dim tgtPerson As Person
    Dim hits As Collection
    Dim tgtPeople As Object 'Dictionary
    Dim tgtSeminar As Variant
    Dim hitsCnt As Long
    Dim altMailFlg As Boolean
    Dim ulFailFlg As Boolean
    Dim msg As String
    Dim topIdx As Long
    
    Set semiList = New SeminarList
    
    If Not semiList.setEvent(iweb.corpName) Then Exit Function
    
    Set tgtPeople = CreateObject("Scripting.Dictionary")
    
    mailInfo.Add vbCrLf & "���Z�~�i�[�C���|�[�g��"
    
    topIdx = mailInfo.Count
    
    For Each tgtSeminar In Array(myNavSeminor, rkNavSeminar)
        If Not tgtSeminar Is Nothing Then
            For Each pid In tgtSeminar.allPeople
                Set tgtPerson = tgtSeminar.allPeople(pid)
                Set hits = iwebPeople.findPerson(tgtPerson)
                
                If hits.Count = 0 Then
                    Set hits = iwebPeople.findPerson(tgtPerson, True)
                    If hits.Count = 1 Then
                        With tgtPerson
                            msg = "��" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "�́Ai-Web���" & hits(1).id & "�ƃ��[���A�h���X���s��v�ł����A�����E��w���E�g�ѓd�b�ԍ�����v�������ߓ���l���Ƃ��Ĉ����܂��B"
                            opeLog.Add msg
                            mailInfo.Add msg
                            msg = vbNullString
                        End With
                    End If
                End If
                
                hitsCnt = hits.Count
                
                If hitsCnt = 1 Then
                    If Not tgtPeople.Exists(hits(1).id) Then
                        tgtPeople.Add hits(1).id, tgtPerson
                    Else
                        tgtPeople(hits(1).id).fusion tgtPerson
                    End If
                ElseIf hits.Count = 0 Then
                    With tgtPerson
                        opeLog.Add "���y���s�z" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "��i-Web�Ɉ�v����l����������܂���ł����B"
                        opeLog.Add "���[���A�h���X�F" & .mailAddress & " / " & .mobileAddress & vbCrLf & "�d�b�ԍ��F" & .mobileNumber & vbCrLf & "��w���F" & .university
                    End With
                    altMailFlg = True
                Else
                    With tgtPerson
                        opeLog.Add "���y���s�z" & (.id & ":" & .kanjiFamilyName & " " & .kanjiFirstName) & "��i-Web�ɏ�񂪈�v����l������������A����ł��܂���B" _
                                    & vbCrLf & "��񂪈�v����l����ID�F" & hits(1).id & " " & "��" & hits.Count - 1 & "��"
                    End With
                    altMailFlg = True
                End If
            Next
        End If
    Next
    
    If altMailFlg Then msg = "�i�r���̌l����i - Web���̌l��񂪈�v������Ώۂ̐l��������ł��Ȃ��o�^�҂����܂��" & vbCrLf
    
    If tgtPeople.Count = 0 Then
        opeLog.Add "�X�V�Ώۂ̃Z�~�i�[���O�͂���܂���ł����B"
        mailInfo.Add "�Ώێ҂����܂���ł����B"
    Else
        'iweb�X�V��������葹�˂��False���Ԃ��Ă���
        If Not semiList.setData(tgtPeople, iweb, myNavi) Then altMailFlg = True
        
        mailInfo.Add semiList.countMember & "�����̃Z�~�i�[�\������X�V���܂����B" & vbCrLf, after:=topIdx
        
        Dim jbName As Variant
        Dim smName As Variant
        Dim ulList As Variant
        
        Set ulList = semiList.outPutList
        
        For Each jbName In ulList
            For Each smName In ulList(jbName)
                '�K���L�����Z������
                '�L�����Z���̓X�e�b�v���P�ʂōs����̂ŁA�����X�e�b�v���̉��ŁA�C�x���gID�@���L�����Z������ɃC�x���gID�A���\��A
                '�ƂȂ����ꍇ�A�\����ɂ��Ă��܂��ƁA��̃L�����Z���̏����i�C�x���gID�����ʁj�ŃC�x���gID�A���L�����Z������Ă��܂��B
                '�t�ɁA�\�񂵂Ă����L�����Z�������ꍇ�A�����C�x���gID�Ȃ�fusion�֐��ŏ㏑�������̂Ŗ�薳���A
                '�ႤID�Ȃ�L�����Z����o�^�ƂȂ邪�A�L�����Z���Ώۂ͈ႤID�Ȃ̂ŁA�Y���C�x���gID�͗\���ԁi�L�����Z������ĂȂ���ԁj�Ŗ��Ȃ�����OK�B
                If Not ULSeminarData(iweb, jbName, smName, ulList(jbName)(smName)(bookState.Cancel), True) Then ulFailFlg = True
                If Not ULSeminarData(iweb, jbName, smName, ulList(jbName)(smName)(bookState.book), False) Then ulFailFlg = True
            Next
        Next
    End If
    
    If ulFailFlg Then
        msg = msg & "�Z�~�i�[���̃A�b�v���[�h�Ɏ��s�����o�^�҂����܂��B" & vbCrLf
        altMailFlg = True
    End If
    
    On Error Resume Next
    loginiWeb iweb
    selectTab iweb, "�S�ĕ\��"
    On Error GoTo 0
    
    If altMailFlg Then
        sendSemAlert msg & "���Y�̓o�^�҂̓Z�~�i�[�̃A�b�v���[�h�����s�ł��Ă���܂���B" & vbCrLf & "�ڍׂ͎��s���O���m�F���Ă��������B"
    Else
        ULAllSeminarData = True
    End If

End Function

Public Function getIwebTopByXMLHTTP(ByVal baseURL As String, ByVal userName As String, ByVal userPass As String) As Object
    Dim objHTTP As Object
    Dim sendData As Variant
    Dim tgtAdd As String
    Dim resTxt As String

    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    With objHTTP
        sendData = "id=" & userName & "&pass=" & userPass
        tgtAdd = "login/check"
    
        .Open "POST", baseURL & tgtAdd, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send sendData
    End With

    If objHTTP.Status = 200 Then
        If InStr(objHTTP.responseText, "���O�C���ł��܂���ł���") Then GoTo loginErr
        Set getIwebTopByXMLHTTP = objHTTP
    Else
        GoTo loginErr
    End If
       
Exit Function
loginErr:
    opeLog.Add "���y���s�z�ʐM�G���[�ɂ��iWeb�Ƀ��O�C���ł��܂���ł����B" & vbCrLf & "�A�h���X/�A�J�E���g/�p�X���[�h/�ʐM�󋵂����m�F���������B"

End Function

Public Function getIwebSeminarStateByXMLHTTP(ByVal baseURL As String, _
                                ByRef objHTTP As Object, _
                                ByVal tgtIwebId As String, _
                                ByVal tgtSeminarNo As String, _
                                ByVal lastUpdate As Date) As Long
    Dim sendData As Variant
    Dim tgtAdd As String
    Dim resTxt As String
    Dim tbl As Variant
    
    '�ŐV�̏󋵂��f�t�H���g�Łu�s���v���i�r���̍X�V�̕����V����
    getIwebSeminarStateByXMLHTTP = bookState.Unknown

    With objHTTP
        sendData = "gkscd=" & tgtIwebId
        tgtAdd = "student/history/"
    
        .Open "POST", baseURL & tgtAdd, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send sendData
    End With

    If objHTTP.Status = 200 Then
        resTxt = objHTTP.responseText
    Else
        GoTo err
    End If
    
    'Dim hitFlg As Boolean
    Dim htmlDoc As Object
    
    Set htmlDoc = New HTMLDocument

    htmlDoc.write resTxt

    '�X�V�����������ꍇ�́A�I���B�ŐV�̏󋵂́u�s���v���i�r���̍X�V�̕����V����
    If InStr(htmlDoc.getElementById("subContents").innerText, "�X�V�����͂���܂���") > 0 Then Exit Function
    
    '���O�\���擾
    Set tbl = htmlDoc.getElementById("selectTable")
    
    '���O�\�������ꍇ�͑z��O�A�G���[
    If tbl Is Nothing Then GoTo err
    
    '�\����Y���Z�~�i�[�̍ŐV�̃��O�i�s���A�\��A�L�����Z���j���擾
    getIwebSeminarStateByXMLHTTP = getLastRecord(tbl, lastUpdate, tgtSeminarNo)

       
Exit Function
err:
    getIwebSeminarStateByXMLHTTP = -1
    opeLog.Add "���y���s�ziWebID : " & tgtIwebId & " �͒ʐM�G���[�ɂ��i-Web�̃Z�~�i�[���O���擾�ł��܂���ł����B"

End Function

Private Function getLastRecord(ByVal logTable As Object, ByVal lastUpdate As Date, ByVal tgtSeminarNo As String) As Long
    Dim i As Long
    Dim n As Long
    Dim MultipleRows As Collection
    Dim logRow As Object
    Dim editBefore As String
    Dim editAfter As String
    Dim seminarJobStep As String
    
    seminarJobStep = SeminarSh.getSminarJobStep(tgtSeminarNo)
    
    Set MultipleRows = New Collection
       
    '���O���Ō�̍s����ǂ�������i�������O��j
    For i = logTable.Rows.Length - 1 To 0 Step -1
        With logTable.Rows(i)
            '�ŏ��̃Z�������t�ŁA
            If IsDate(.Cells(0).innerText) Then
                If CDate(.Cells(0).innerText) < lastUpdate Then
                    '�i�r���ŐV�������Â���ΏI��
                    '�ŐV�̏󋵂́u�s���v���i�r���̍X�V�̕����V����
                    getLastRecord = bookState.Unknown
                    Exit Function
                Else
                    '�i�r���ŐV�������V�����ꍇ�͕ێ�
                    MultipleRows.Add logTable.Rows(i)
                End If
            Else
            '�ŏ��̃Z�������t�łȂ��Ȃ畡���s�̃p�^�[���Ȃ̂ł�������ێ����Ď��̍s�ֈړ�
                MultipleRows.Add logTable.Rows(i)
                GoTo CONTINUE
            End If
        End With
        
        '�ŏ��̃Z�������t�ŁA�i�r���ŐV�������V�����ꍇ
        
        For Each logRow In MultipleRows
            With logRow
            '�Z�������V�̏ꍇ�ƂS�̏ꍇ�őΏۂƂ���Z����ς���
            n = .Cells.Length
            
            '���O�e�L�X�g����A�hi-Web�C�x���gID�h���h�L�����Z���h�𒊏o���Ă���B
                If n = 7 Then
                    editBefore = getEventIDfromLogText(.Cells(n - 4).innerText)
                    editAfter = getEventIDfromLogText(.Cells(n - 3).innerText)
                    
                ElseIf n = 4 Then
                    editBefore = getEventIDfromLogText(.Cells(n - 3).innerText)
                    editAfter = getEventIDfromLogText(.Cells(n - 2).innerText)
                End If
            End With
            
            '�ύX��ɋL�ڂ�����Ƃ��i�Ȃ����̓X�L�b�v�j
            If editAfter <> vbNullString Then
                If editAfter = "�L�����Z��" Then
                    '�ύX�オ�L�����Z���ŁA�ύX�O�̐E��{�X�e�b�v���Y���Ɠ������ꍇ�A�ŐV�̏󋵂́u�L�����Z���v
                    '�����I��
                    If SeminarSh.getSminarJobStep(editBefore) = seminarJobStep Then
                         getLastRecord = bookState.Cancel
                         Exit Function
                    End If
                ElseIf SeminarSh.getSminarJobStep(editAfter) = seminarJobStep Then
                    '�ύX�オ�L�����Z���łȂ��A���ύX�O�̐E��{�X�e�b�v���Y���Ɠ������ꍇ�A�ŐV�̏󋵂́u�\��v
                    '�����I��
                    getLastRecord = bookState.book
                    Exit Function
                End If
            End If
        Next
        
        '�ێ����Ă����s���N���A�ɂ���B
        Set MultipleRows = New Collection
CONTINUE:
    Next
    
    '�S���O�ɊY�����Ȃ���΁A�ŐV�̏󋵂́u�s���v���i�r���̍X�V�̕����V����
    getLastRecord = bookState.Unknown
    
End Function

Private Function getEventIDfromLogText(ByVal LogText As String) As String
    Dim reg As Object
    Dim match As Object
    
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = "No\.?(\d+)[ :]|(�L�����Z��)"
    End With
    
    For Each match In reg.Execute(LogText)
        getEventIDfromLogText = IIf(IsEmpty(match.SubMatches(0)), match.SubMatches(1), match.SubMatches(0))
    Next
End Function
