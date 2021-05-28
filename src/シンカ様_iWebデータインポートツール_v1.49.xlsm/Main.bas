Attribute VB_Name = "Main"
Option Explicit

#Const cnsTest = 0   '#�{��
'#Const cnsTest = 1     '#�e�X�g
'�{��/�e�X�g��؂�ւ���ꍇ�́A�A�J�E���g�V�[�g�̒萔�����������邱�ƁI

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const LOCK_PSWD_RNG As String = "B3"

Private Const CAP_SUB_RNG As String = "E4"
Private Const FIN_SUB_RNG As String = "E11"

Public Const AC_CORP_IDX As Long = 1
Public Const AC_SITE_IDX As Long = 2
Public Const AC_ADDR_IDX As Long = 3
Public Const AC_COID_IDX As Long = 4
Public Const AC_ACNT_IDX As Long = 5
Public Const AC_PSWD_IDX As Long = 6
Public Const AC_DLLO_IDX As Long = 7
Public Const AC_EDLO_IDX As Long = 8
Public Const AC_IWUL_IDX As Long = 9
Public Const AC_IWDL_IDX As Long = 10
Public Const AC_IWCO_IDX As Long = 11

Public Const SC_CORP_IDX As Long = 2
Public Const SC_MLFL_IDX As Long = 8
Public Const SC_LAST_UPDT_COL_IDX As Long = 11

Public Const NO_USER_MSG As String = "(���҂Ȃ�)"

Public Const OUT_DATA_ERR As Long = 1 + vbObjectError + 512
Public Const GET_CUR_REG_ERR As Long = 2 + vbObjectError + 512

Public opeLog As Collection
Public mailInfo As Collection
Public cancelFlg As Boolean

Enum dataType
    personal = 1
    Seminar = 2
End Enum

Sub Main()
    Dim scRng As Range '�u���C���v�V�[�g�\�J�n�Z��(�u���ԁv)
    Dim initTime As Date
    Dim finMailMsg As String
    Dim i As Long
    
    '3�̑O����
    '�@�u���s�Ҏ����v�̃f�[�^�`�F�b�N�A�A�ΏۃV�[�g�̕ی�`�F�b�N�A�B�ΏۃV�[�g�̃t�B���^�[����
    If Not preCheck Then
        Exit Sub
    End If
    
    cancelFlg = False
    initTime = Now()
      
    Set opeLog = New Collection  '�R���N�V�����g�p

    '�u���C���V�[�g�v�̕\�f�[�^���擾
    Set scRng = getCurrentRegion(ScenarioSh.Cells(2, 1), 1, False)
    
    On Error Resume Next
    '�u���s���O�v�V�[�g�̕\�f�[�^���擾���A�Z���̒l���N���A
    getCurrentRegion(LogSh.Cells(2, 1), 1, False).ClearContents
    On Error GoTo 0
    
'###### 20200602 AIM�FYamamoto ######
'�Ώۊ�Ɩ��Ɏg�p�֎~�������Ȃ�������
    If corpNameCheck(scRng) Then
#If cnsTest = 1 Then
        MsgBox "�Ώۊ�Ɩ��Ɏg�p�֎~�����Ȃ�"
        Exit Sub
#End If
    Else
        MsgBox "�Ώۊ�Ɩ��Ɏg�p�֎~�������܂܂�Ă��܂��B"
        Exit Sub
    End If
'####################################
    
    For i = 1 To scRng.Rows.Count  '10�s�ڂ܂�(���Ԃ�10�����͂���Ă���)
        '�������[����������
        Set mailInfo = New Collection  '�R���N�V�����g�p
        
        ' �u���C���v�V�[�g�̍��ږ��u���s�v��TRUE������
        '�y�v�ύX�z��ԍ��͒萔�ɕύX�����ق����ǂ�����
        If scRng(i, 7).Value Then
            'TRUE�̏ꍇ�@�i���L�ǂ�����A�uWorksheet_Change�v�����s�����j
            scRng(i, 9).Value = initTime  '���ږ��u�J�n�����v�ɏ�L�Ŏ擾����Now�����
            scRng(i, 10).Value = vbNullString  '���ږ��u�������ʁv���󗓂ɂ���

            If excProcess(scRng(i, 2).Value) Then
                scRng(i, 10).Value = "OK"
                
                If ScenarioSh.getMyNaviOmit Or ScenarioSh.getRikuNaviOmit Or ScenarioSh.getPsOmit Or ScenarioSh.getSmOmit Then
                    opeLog.Add "�����̖�����TRUE�̂��߁A�I�������͍X�V���܂���B" & vbCrLf & _
                               "����̍X�V�����F" & initTime
                    scRng(i, 11).offset(0, 1).Value = "*"
                Else
                    scRng(i, 11).Value = initTime
                End If
                
                '�X�V���������t�@�C�����X�V
                SettingSh.ensureOldPath scRng(i, 2).Value
                
                finMailMsg = scRng(i, 2).Value & "��i-Web�C���|�[�g���������܂����I" & vbCrLf & vbCrLf & getMailInfo()
                
                If sendFinAlert(finMailMsg, scRng(i, 2).Value) Then
                    opeLog.Add scRng(i, 2).Value & "��i-Web�C���|�[�g���������܂����I"
                    outputLog "Import Completed.", True, scRng(i, 2).Value, vbNullString
                Else
                    opeLog.Add "�����̃��[���ʒm�������Ă���܂���B"
                    outputLog "Finish Alert Mail could not send.", True, scRng(i, 2).Value, vbNullString
                End If
            Else
                scRng(i, 10).Value = "NG"
            End If
            
        End If
    Next
    
    If Not LogSh.AutoFilterMode Then LogSh.setAutoFilter
    If Not OldLogSh.AutoFilterMode Then OldLogSh.setAutoFilter

    Unload AlertBox
    Set opeLog = Nothing
    
End Sub

'3�̑O����
'�@���s�Ҏ����̃f�[�^�L���`�F�b�N�A�A�V�[�g�ی�`�F�b�N�A�B�t�B���^�[�`�F�b�N
Public Function preCheck() As Boolean
    Dim sh As Variant
    Dim name As String
    
    name = ScenarioSh.getUserName
    
    '�u���C���v�V�[�g�́u���s�Ҏ����v�f�[�^�L���`�F�b�N
    If name = vbNullString Or name = NO_USER_MSG Then
        MsgBox "���s�Җ����󔒂ł��B" & vbCrLf & "���s�Җ���I�������̂��A�ēx���s���Ă��������B"
        Exit Function
    End If

    '4�̃V�[�g�̃V�[�g�ی�`�F�b�N(�@�u���C���v�A�A�u�A�J�E���g�v�A�B�u�ߋ����O�v�A�C�u���[���A�J�E���g�v�V�[�g)
    '�V�[�g�ی삪�|�����Ă��Ȃ��ꍇ�̓A���[�g�\��
    For Each sh In Array(ScenarioSh, AccountSh, OldLogSh, MailSettingSh)
        If Not sh.ProtectContents Then
            MsgBox sh.name & "�V�[�g���ی�������ł��B" & vbCrLf & "�ی���ĊJ�����̂��A�ēx���s���Ă��������B"
            Exit Function
        End If
    Next
    
    '2�̃V�[�g�Ƀt�B���^�[���|�����Ă���ꍇ�̓t�B���^�[����(�@�u���s���O�v�A�A�u�ߋ����O�v�V�[�g)
    If LogSh.AutoFilterMode Then LogSh.setAutoFilter
    If OldLogSh.AutoFilterMode Then OldLogSh.setAutoFilter
    
    preCheck = True

End Function

'##### 20200602 AIM�FYamamoto #####
'�Ώۊ�Ɩ��Ɏg�p�֎~�������܂܂�Ă����ꍇ�A���O�ɒǋL
Private Function corpNameCheck(scRng As Range) As Boolean
    Dim i As Long
    Dim re As New RegExp
    
    corpNameCheck = True

    For i = 1 To scRng.Rows.Count
        With re
            .Global = True
            .Pattern = "([!#$%&'""`+\-/=~,;:@^<>\?\*\|\{\}\(\)\[\]\\])+"
            If .test(scRng.Cells(i, 2).Value) Then
                corpNameCheck = False
                opeLog.Add "�Ώۊ�Ɩ��Ɏg�p�֎~�������܂܂�Ă��܂��B"
                outputLog "Corporate Name Check ", False, scRng.Cells(i, 2).Value, vbNullString
            End If
        End With
    Next
End Function

Private Function excProcess(ByVal tgtCorpName As String) As Boolean

    AlertBox.Caption = tgtCorpName & "���s��"
    AlertBox.Show False

'###IE���b�p�[���N��
    '##i-Web�p��IE���b�p�[���N��
    Dim iweb As CorpSite
    Set iweb = New CorpSite
    
    On Error GoTo wrapperErr
    '���O�ǉ�
    opeLog.Add "InternetExplore���N����..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr
    iweb.cleanUpTgtSite

    '##�}�C�i�r�p��IE���b�p�[���N��
    Dim myNavi As CorpSite
    Dim myNaviFlg As Boolean
    Set myNavi = New CorpSite

    If myNavi Is Nothing Then GoTo wrapperErr
    If myNavi.setCorp(tgtCorpName, "�}�C�i�r") Then
        myNaviFlg = True
        myNavi.cleanUpTgtSite
    End If
      
    '##���N�i�r�p��IE���b�p�[���N��
    Dim rikuNavi As CorpSite
    Dim rikuNaviFlg As Boolean
    Set rikuNavi = New CorpSite

    If rikuNavi Is Nothing Then GoTo wrapperErr
    If rikuNavi.setCorp(tgtCorpName, "���N�i�r") Then
        rikuNaviFlg = True
        rikuNavi.cleanUpTgtSite
    End If
    
    '##���C���V�[�g�̃}�C�i�r�����A���N�i�r�����𔽉f������
    If ScenarioSh.getMyNaviOmit Then myNaviFlg = False
    If ScenarioSh.getRikuNaviOmit Then rikuNaviFlg = False
    
    If Not myNaviFlg And Not rikuNaviFlg Then
        opeLog.Add "���s�\�ȃA�J�E���g���Ȃ����ߏ������I�����܂��B"
        GoTo wrapperErr
    End If
    
    Dim psFlg As Boolean: psFlg = True
    Dim smFlg As Boolean: smFlg = True
    
    '##���C���V�[�g�̃}�C�i�r�����A���N�i�r�����𔽉f������
    If ScenarioSh.getPsOmit Then psFlg = False
    If ScenarioSh.getSmOmit Then smFlg = False
    
    If Not psFlg And Not smFlg Then
        opeLog.Add "���s�\�ȏ������Ȃ����ߏ������I�����܂��B"
        GoTo wrapperErr
    End If
        
    '���O�o��
    opeLog.Add "�N�������B"
    outputLog "InternetExplore Wrapper startup", True, tgtCorpName, vbNullString
    On Error GoTo 0

'####i-web����l����S���_�E�����[�h
    Dim iWebPsDataFilePath As String
    
    If psFlg Then
        opeLog.Add "i-web����l����S���_�E�����[�h�J�n"
    
        On Error GoTo dlErr
#If cnsTest = 0 Then
        iWebPsDataFilePath = DLiWebData(iweb, iweb.iWebPDLayNo)

        iWebPsDataFilePath = moveFileAddHeadder(iWebPsDataFilePath, "�y" & tgtCorpName & "�z" & "i-Web�l���S��_�l���UL�O_")

        If iWebPsDataFilePath = vbNullString Then GoTo dlErr
#End If
        opeLog.Add "�S���_�E�����[�h����"
    
        On Error GoTo 0
        
    Else
        opeLog.Add "�l���C���|�[�g�����̂��߁Ai-web����l����S���_�E�����[�h���X�L�b�v���܂��B"
    
    End If
    
    outputLog "Download i-Web personal data", True, tgtCorpName, iWebPsDataFilePath
    

'###�}�C�i�r����l���/�Z�~�i�[�����_�E�����[�h
    Dim myNavPsDataFilePath As String
    Dim myNavPsDataFileName As String
    Dim myNavPsFlg As Boolean
    Dim mynavSmDataFilePath As String
    Dim myNavSmDataFileName As String
    Dim myNavSmFlg As Boolean

    '2018/07 �ȍ~�e�X�g�ł̃}�C�i�r���O�C���͋֎~�I�K�v�ł���΂��q�l�̋��𓾂鎖�I

    If Not myNaviFlg Then
        opeLog.Add "�}�C�i�r�̏����͂���܂���B"
        outputLog "Skip Mynavi Download/Upload", True, tgtCorpName, vbNullString
        GoTo MY_NAV_SKIP
    End If

#If cnsTest = 1 Then
    '#�e�X�g�p�_�~�[�f�[�^
    If Not psFlg Then
        opeLog.Add "�l���C���|�[�g�����̂��߁A�}�C�i�r����̌l���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
    Else
        myNavPsFlg = True
        myNavPsDataFilePath = getDlFilePath("20000�V���J_�}�C_�l.csv")
    End If
    
    If Not smFlg Then
        opeLog.Add "�Z�~�i�[�C���|�[�g�����̂��߁A�}�C�i�r����̃Z�~�i�[���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip MyNavi Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        myNavSmFlg = True
        'mynavSmDataFilePath = getDlFilePath("20000�V���J_�Z�~�i_�l.csv")
        mynavSmDataFilePath = "C:\Users\11402086\Desktop\test\2\20000�V���J_�}�C_�Z�~�i2.csv"
        SettingSh.Cells(11, 2).Value = "C:\Users\11402086\Desktop\test\2\20000�V���J_�}�C_�Z�~�i.csv"
    End If
       
    GoTo MY_NAV_SKIP
#End If

    '##�}�C�i�r���O�C��
    AlertBox.Label1 = "�}�C�i�r�Ƀ��O�C����.."
    opeLog.Add "�}�C�i�r�Ƀ��O�C����.. �A�J�E���g�F" & myNavi.userName

    On Error GoTo loginErr
    If Not loginMyNavi(myNavi) Then GoTo loginErr
    On Error GoTo 0

    opeLog.Add "�}�C�i�r�Ƀ��O�C������"
        
    '##�}�C�i�r�l���DL�\��
    If Not psFlg Then
        opeLog.Add "�l���C���|�[�g�����̂��߁A�}�C�i�r����̌l���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
    Else
        '##�X�V������ݒ肵�Č���
        AlertBox.Label1 = "�X�V������ݒ肵�Č�����.."

        opeLog.Add "�}�C�i�r�̌l��񌟍��y�[�W�Ɉړ���.."
        On Error GoTo pageErr
        If Not moveSearchWindow(myNavi) Then GoTo pageErr
        On Error GoTo 0

        opeLog.Add "�ړ������B�l��񌟍��J�n"
        
        myNavPsFlg = myNavi.dlLayout <> vbNullString
        
        If Not myNavPsFlg Then
            opeLog.Add "�i�r�T�C�g���C�A�E�g���i�l���Download�p�j���A�J�E���g�V�[�g�ɋL�ڂ���Ă��܂���B�X�L�b�v���܂��B"
            outputLog "Skip MyNavi Personal Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchMyNaviDt(myNavi, dataType.personal) Then
                opeLog.Add "�V�K�f�[�^�Ȃ��B"
                outputLog "Did not hit new MyNavi Personal Data", True, tgtCorpName, vbNullString
                
                myNavPsFlg = False
            Else
                '#CSV�t�@�C��DL��\��
                AlertBox.Label1 = "�}�C�i�r�Ōl���CSV���쐬��.."
                opeLog.Add "�}�C�i�r�Ōl���CSV���쐬��.."
    
                myNavPsDataFileName = makeMyNaviDt(myNavi, dataType.personal)
                If myNavPsDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
    End If
        
    '##�}�C�i�r�Z�~�i�[���DL�\��
        
    myNavSmFlg = False
        
    If Not smFlg Then
        opeLog.Add "�Z�~�i�[�C���|�[�g�����̂��߁A�}�C�i�r����̃Z�~�i�[���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip MyNavi Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# �C�x���g����S����
        AlertBox.Label1 = "�Z�~�i�[����S������.."
        opeLog.Add "�}�C�i�r�ŃZ�~�i�[���̌����J�n.."
        
        On Error GoTo dlErr

        myNavSmFlg = myNavi.dlLayoutEV <> vbNullString

        If Not myNavSmFlg Then
            opeLog.Add "�i�r�T�C�g���C�A�E�g���i�C�x���g�E�Z�~�i�[���Download�p�j���A�J�E���g�V�[�g�ɋL�ڂ���Ă��܂���B�X�L�b�v���܂��B"
            outputLog "Skip MyNavi Seminar Data Download ", True, tgtCorpName, vbNullString
        Else
            opeLog.Add "�}�C�i�r�̃Z�~�i�[��񌟍��y�[�W�Ɉړ���.."

            On Error GoTo pageErr
            If Not moveSearchWindow(myNavi, Not psFlg) Then GoTo pageErr
            On Error GoTo 0

            If Not searchMyNaviDt(myNavi, dataType.Seminar) Then
                If InStr(opeLog(opeLog.Count), "�f�[�^�͂���܂���ł����B") = 0 Then
                    GoTo dlErr
                Else
                    opeLog.Add "�V�K�f�[�^�Ȃ��B"
                    outputLog "Did not hit new MyNavi Seminar Data", True, tgtCorpName, vbNullString
                    myNavSmFlg = False
                End If
            Else
                '#CSV�t�@�C��DL��\��
                AlertBox.Label1 = "�}�C�i�r�ŃZ�~�i�[���CSV���쐬��.."
                opeLog.Add "�}�C�i�r�ŃZ�~�i�[���CSV���쐬��.."

                myNavSmDataFileName = makeMyNaviDt(myNavi, dataType.Seminar)
                If myNavSmDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
        
        On Error GoTo 0
    End If
    
    '# CSV�t�@�C�����_�E�����[�h

    If myNavPsFlg Then chkDateCreated myNavi, myNavPsDataFileName
    If myNavSmFlg Then chkDateCreated myNavi, myNavSmDataFileName
    
    If myNavPsFlg Then
        opeLog.Add "�l���f�[�^�_�E�����[�h�J�n.."
        myNavPsDataFilePath = dlMyNaviCSV(myNavi, myNavPsDataFileName)
        If myNavPsDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "�l���f�[�^�_�E�����[�h�����B"
        outputLog "Download myNavi personal CSV data", True, tgtCorpName, myNavPsDataFilePath
    End If
    
    If myNavSmFlg Then
        opeLog.Add "�Z�~�i�[�f�[�^�_�E�����[�h�J�n�B"
        mynavSmDataFilePath = dlMyNaviCSV(myNavi, myNavSmDataFileName)
        If mynavSmDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "�Z�~�i�[�f�[�^�_�E�����[�h�����B"
        outputLog "Download myNavi Seminar CSV data", True, tgtCorpName, mynavSmDataFilePath
    End If

MY_NAV_SKIP:

' ���ǉ���������������������������������
    If myNavSmFlg Then
        mynavSmDataFilePath = getDiffFile(mynavSmDataFilePath, tgtCorpName, myNavi.lastUpdate)
        If mynavSmDataFilePath = vbNullString Then GoTo diffErr
    End If
' ��������������������������������������

'###�}�C�i�r�T�C�g�̕\���I��
    On Error Resume Next
    myNavi.visible False
    On Error GoTo 0

'###���N�i�r����l���/�Z�~�i�[�����_�E�����[�h
    Dim rkNavPsDataFilePath As String
    Dim rkNavPsDataFileName As String
    Dim rkNavPsFlg As Boolean
    Dim rknavSmDataFilePath As String
    Dim rkNavSmDataFileName As String
    Dim rkNavSmFlg As Boolean

    If Not rikuNaviFlg Then
        opeLog.Add "���N�i�r�̏����͂���܂���B"
        outputLog "Skip RikuNavi Download/Upload", True, tgtCorpName, vbNullString
        GoTo RIKU_NAV_SKIP
    End If

#If cnsTest = 1 Then
    If Not psFlg Then
        opeLog.Add "�l���C���|�[�g�����̂��߁A���N�i�r����̌l���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip RikuNaviSeminar CSV data ", True, tgtCorpName, vbNullString
    Else
        rkNavPsFlg = True
        rkNavPsDataFilePath = getDlFilePath("20000�V���J_���N_�l.csv")
    End If
    
    If Not smFlg Then
        opeLog.Add "�Z�~�i�[�C���|�[�g�����̂��߁A���N�i�r����̃Z�~�i�[���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip RikuNavi  Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        rkNavSmFlg = True
        'rknavSmDataFilePath = getDlFilePath("20000�V���J_���N_�Z�~�i.csv")
        rknavSmDataFilePath = "C:\Users\11402086\Desktop\test\2\20000�V���J_���N_�Z�~�i.csv"
    End If
        
    GoTo RIKU_NAV_SKIP
#End If
    
    
    '# ���N�i�r���O�C��
    AlertBox.Label1 = "���N�i�r�Ƀ��O�C����.."
    opeLog.Add "���N�i�r�Ƀ��O�C����.. �A�J�E���g�F" & rikuNavi.userName

    On Error GoTo loginErr
    If Not loginRikuNavi(rikuNavi) Then GoTo loginErr
    On Error GoTo 0

    opeLog.Add "���N�i�r�Ƀ��O�C������"

    '##���N�i�r�l���DL�\��
    If Not psFlg Then
        opeLog.Add "�l�C���|�[�g�����̂��߁A���N�i�r����̌l���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip RikuNaviSeminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# �X�V������ݒ肵�ĐV�K�o�^���ꂽ�l��񌟍�
        AlertBox.Label1 = "�X�V������ݒ肵�Č�����.."
        opeLog.Add "���N�i�r�Ōl���̌����J�n.."

        On Error GoTo dlErr

        rkNavPsFlg = rikuNavi.dlLayout <> vbNullString

        If Not rkNavPsFlg Then
            opeLog.Add "�i�r�T�C�g���C�A�E�g���i�l���Download�p�j���A�J�E���g�V�[�g�ɋL�ڂ���Ă��܂���B�X�L�b�v���܂��B"
            outputLog "Skip RikuNavi Personal Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchRikuNaviDt(rikuNavi, dataType.personal) Then
                If InStr(opeLog(opeLog.Count), "�Ō������܂������Y������f�[�^�͂���܂���ł����B") = 0 Then
                    GoTo dlErr
                Else
                    opeLog.Add "�l���̐V�K�o�^�Ȃ��B"
                    outputLog "RikuNavi : Did not hit new Personal Data", True, tgtCorpName, vbNullString
    
                    rkNavPsFlg = False
                End If
            Else
                rkNavPsDataFileName = makeRikuNaviDt(rikuNavi, dataType.personal)
                If rkNavPsDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
        
        On Error GoTo 0
    End If

    '##���N�i�r�Z�~�i�[���DL�\��
    rkNavSmFlg = False
    
    If Not smFlg Then
        opeLog.Add "�Z�~�i�[�C���|�[�g�����̂��߁A���N�i�r����̃Z�~�i�[���_�E�����[�h���X�L�b�v���܂��B"
        outputLog "Skip RikuNavi  Seminar CSV data ", True, tgtCorpName, vbNullString
    Else
        '# �C�x���g����S����
        AlertBox.Label1 = "�Z�~�i�[����S������.."
        opeLog.Add "���N�i�r�ŃZ�~�i�[���̌����J�n.."

        rkNavSmFlg = rikuNavi.dlLayoutEV <> vbNullString

        If Not rkNavSmFlg Then
            opeLog.Add "�i�r�T�C�g���C�A�E�g���i�C�x���g�E�Z�~�i�[���Download�p�j���A�J�E���g�V�[�g�ɋL�ڂ���Ă��܂���B�X�L�b�v���܂��B"
            outputLog "Skip RikuNavi Seminar Data Download ", True, tgtCorpName, vbNullString
        Else
            If Not searchRikuNaviDt(rikuNavi, dataType.Seminar) Then
                opeLog.Add "�Z�~�i�[���̓o�^�Ȃ��B"
                outputLog "RikuNavi : Did not hit new Seminar Data", True, tgtCorpName, vbNullString

                rkNavSmFlg = False
            Else
                rkNavSmDataFileName = makeRikuNaviDt(rikuNavi, dataType.Seminar)
                If rkNavSmDataFileName = vbNullString Then GoTo csvErr
            End If
        End If
    End If

    '# CSV�̍쐬��ҋ@
    If rkNavPsFlg Then
        AlertBox.Label1 = "���N�i�r�Ōl���p��CSV���쐬���i���Ԃ�������܂��j.."
        opeLog.Add "���N�i�r�Ōl���p��CSV���쐬��.."
        waitRikuNaviCSV rikuNavi, rkNavPsDataFileName
    End If

    If rkNavSmFlg Then
        AlertBox.Label1 = "���N�i�r�ŃZ�~�i�[���p��CSV���쐬���i���Ԃ�������܂��j.."
        opeLog.Add "���N�i�r�ŃZ�~�i�[���p��CSV���쐬��.."
        waitRikuNaviCSV rikuNavi, rkNavSmDataFileName
    End If

    '# CSV�t�@�C�����o���������DL���J�n�ADL����������p�X���Ԃ��Ă���
    If rkNavPsFlg Then
        opeLog.Add "�l���_�E�����[�h�J�n.."
        rkNavPsDataFilePath = dlRikuNaviCSV(rikuNavi, rkNavPsDataFileName)
        If rkNavPsDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "�l���_�E�����[�h�����B"
        outputLog "Download RikuNavi personal data", True, tgtCorpName, rkNavPsDataFilePath
    End If

    If rkNavSmFlg Then
        opeLog.Add "�Z�~�i�[���_�E�����[�h�J�n.."
        rknavSmDataFilePath = dlRikuNaviCSV(rikuNavi, rkNavSmDataFileName)
        If rknavSmDataFilePath = vbNullString Then GoTo dlErr

        opeLog.Add "�Z�~�i�[���_�E�����[�h�����B"
        outputLog "Download RikuNavi seminar data", True, tgtCorpName, rknavSmDataFilePath
    End If

'###���N�i�r�T�C�g�̕\���I��
    On Error Resume Next
    rikuNavi.visible False
    On Error GoTo 0

RIKU_NAV_SKIP:

    On Error GoTo ulErr
'###�}�C�i�r�̏���i-Web�փA�b�v���[�h
    If Not myNavPsDataFilePath = vbNullString And myNavPsFlg And psFlg Then

        AlertBox.Label1 = "�}�C�i�r��CSV����AiWeb�֌l�����A�b�v���[�h���Ă��܂��B"
        opeLog.Add "�}�C�i�r��CSV����AiWeb�֌l�����A�b�v���[�h��.."
        mailInfo.Add "���l���C���|�[�g��" & vbCrLf & "(�}�C�i�r)"

        If Not ULPersonalData(iweb, myNavPsDataFilePath, myNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "�A�b�v���[�h����"
        outputLog "MyNavi personal data Upload to i-Web", True, tgtCorpName, myNavPsDataFilePath
    Else
        mailInfo.Add "���l���C���|�[�g��" & vbCrLf & "(�}�C�i�r)" & vbCrLf & "�Ώێ҂����܂���ł����B"
    End If

'###���N�i�r�̏���i-Web�փA�b�v���[�h
    If Not rkNavPsDataFilePath = vbNullString And rkNavPsFlg And psFlg Then

        AlertBox.Label1 = "���N�i�r��CSV����AiWeb�֌l�����A�b�v���[�h���Ă��܂��B"
        opeLog.Add "���N�i�r��CSV����AiWeb�֌l�����A�b�v���[�h��.."
        mailInfo.Add "(���N�i�r)"

        If Not ULPersonalData(iweb, rkNavPsDataFilePath, rikuNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "�A�b�v���[�h����"
        outputLog "RikuNavi personal data Upload to i-Web", True, tgtCorpName, rkNavPsDataFilePath
    Else
        mailInfo.Add "(���N�i�r)" & vbCrLf & "�Ώێ҂����܂���ł����B"
    End If

    On Error GoTo 0
      
    If (myNavSmFlg Or rkNavSmFlg) And smFlg Then
    '####i-web����l����S���_�E�����[�h(�ǉ����𔽉f)
        'Dim iWebPsDataFilePath As String
        
        AlertBox.Label1 = "i-web����l����S���_�E�����[�h�J�n(�ǉ����𔽉f).."
        opeLog.Add "i-web����l����S���_�E�����[�h�J�n"

        On Error GoTo dlErr
#If cnsTest = 0 Then
        iWebPsDataFilePath = DLiWebData(iweb, iweb.iWebPDLayNo)

        iWebPsDataFilePath = moveFileAddHeadder(iWebPsDataFilePath, "�y" & tgtCorpName & "�z" & "i-Web�l���S��_�l���UL��_")

        If iWebPsDataFilePath = vbNullString Then GoTo dlErr
#End If
        opeLog.Add "�S���_�E�����[�h����"
        outputLog "Download i-Web personal data", True, tgtCorpName, iWebPsDataFilePath
    
        On Error GoTo 0
        
#If cnsTest = 1 Then
    iWebPsDataFilePath = "C:\Users\11402086\Desktop\test\iwebData.csv"
#End If
       
    '####�t�@�C������i-Web�̌l�������[�h
        Dim iwebPeople As people
    
        Set iwebPeople = getPeople("i-Web", iWebPsDataFilePath)
        
        AlertBox.Label1 = "�t�@�C������i-Web�̌l�������[�h��.."
    
        If iwebPeople Is Nothing Then
            outputLog "Load i-Web personal data from csv file", False, tgtCorpName, iWebPsDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load i-Web personal data from csv file", True, tgtCorpName, iWebPsDataFilePath
        
        End If
    Else
        mailInfo.Add vbCrLf & "���Z�~�i�[�C���|�[�g��" & vbCrLf & "�Ώێ҂����܂���ł����B"
    End If

'####�t�@�C������}�C�i�r�̃Z�~�i�[�������[�h
    Dim myNavSeminor As people
    
     AlertBox.Label1 = "�t�@�C������}�C�i�r�̃Z�~�i�[�������[�h��.."
    
    If Not mynavSmDataFilePath = vbNullString And myNavSmFlg And smFlg Then
        Set myNavSeminor = getPeople("�}�C�i�r", mynavSmDataFilePath, False, iweb.lastUpdate)
    
        If myNavSeminor Is Nothing Then
            outputLog "Load MyNavi seminar data from csv file", False, tgtCorpName, mynavSmDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load MyNavi seminar data from csv file", True, tgtCorpName, mynavSmDataFilePath
        End If
    End If


'####�t�@�C�����烊�N�i�r�̃Z�~�i�[�������[�h
    Dim rkNavSeminar As people
    
    AlertBox.Label1 = "�t�@�C�����烊�N�i�r�̃Z�~�i�[�������[�h��.."
    
    If Not rknavSmDataFilePath = vbNullString And rkNavSmFlg And smFlg Then
        Set rkNavSeminar = getPeople("���N�i�r", rknavSmDataFilePath, False, iweb.lastUpdate)
    
        If rkNavSeminar Is Nothing Then
            outputLog "Load RikuNavi seminar data from csv file", False, tgtCorpName, rknavSmDataFilePath
            excProcess = False
            GoTo normalFin
        Else
            outputLog "Load RikuNavi seminar data from csv file", True, tgtCorpName, rknavSmDataFilePath
        End If
    End If
    
'###�Z�~�i�[���A�b�v���[�h
    
    If (myNavSmFlg Or rkNavSmFlg) And smFlg Then
        AlertBox.Label1 = "�Z�~�i�[�����A�b�v���[�h��.."
    
        If Not ULAllSeminarData(iweb, myNavi, myNavSeminor, rkNavSeminar, iwebPeople) Then GoTo ulErr
        outputLog "Uupload all Seminar data", True, tgtCorpName, vbNullString
    Else
        'Do nothing
    End If
    
'###��������
    excProcess = True

normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not rikuNavi Is Nothing Then rikuNavi.quitAll
    If Not myNavi Is Nothing Then myNavi.quitAll
    
Exit Function

'###�G���[����
wrapperErr:
    outputLog "IE Wrapper could not startup", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin

loginErr:
    outputLog "Login failed", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin

pageErr:
    outputLog "Failed to navigate the target page", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
csvErr:
    outputLog "CSV could not be created on Navi site", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
dlErr:
    outputLog "Failed to download the CSV file", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
diffErr:
    outputLog "Failed to extract difference", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
ulErr:
    outputLog "Failed to upload the data", False, tgtCorpName, vbNullString
    excProcess = False
    GoTo normalFin
    
End Function

Private Function outputLog(ByVal opeName As String, ByVal successFlg As Boolean, ByVal tgtCorpName As String, Optional ByVal tgtFilePath As String)
    Dim tgtCell As Range
    Dim result As String
    Dim errLog As Variant
    Dim errLogs As Collection
    Dim i As Long
    Dim j As Long
    Dim maxLine As Long
    Dim userName As String
    Dim nextFlg As Boolean
    
    Set errLogs = New Collection
    
    If successFlg Then
        result = "����"
    Else
        result = "���s"
    End If
    
    maxLine = SettingSh.getLogMaxLow
    
    If maxLine = 0 Or maxLine > 253 Then maxLine = 253
    
    j = 1
    
    '1���O��32,767�����𒴂�����̂͂Ȃ��i�Ӑ}�I�ɍ��Ȃ��Ƃł��Ȃ��j
    For i = 1 To opeLog.Count
        '���O���܂Ƃ߂�
        errLog = errLog & IIf(j > 1, vbCrLf, vbNullString) & opeLog(i)
        j = j + 1
        
        '���O���w��s�𒴂���A�������͍ŏI���O�̂Ƃ��t���O�𗧂Ă�B
        If j = maxLine + 1 Or i = opeLog.Count Then
            nextFlg = True
        
        '���O���ŏI���O�łȂ��A���̃��O�ƍ��킹��32,767�����𒴂���Ƃ��t���O�𗧂Ă�B
        ElseIf Len(errLog) + Len(opeLog(i + 1)) + 2 > 32767 Then
            nextFlg = True
        End If
        
        '�t���O�������Ă�����A���O's�ɒǉ����āA��������N���A�B
        If nextFlg Then
            errLogs.Add errLog
            errLog = vbNullString
            j = 1
            nextFlg = False
        End If
    Next
    
    'opeLog �N���A
    Set opeLog = New Collection
    
    userName = ScenarioSh.getUserName
    
    For Each errLog In errLogs
        With LogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0)
            .Value = Now()
            .offset(0, 1).Value = result
            .offset(0, 2).Value = tgtCorpName
            .offset(0, 3).Value = opeName
            .offset(0, 4).Value = tgtFilePath
            .offset(0, 5).Value = errLog
            .offset(0, 6).Value = userName
            'OldLogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0).Resize(1, 6).Interior.Color = .Resize(1, 6).Interior.Color
            OldLogSh.Cells(LogSh.Rows.Count, 1).End(xlUp).offset(1, 0).Resize(1, 7).Value = .Resize(1, 7).Value
        End With
    Next

End Function

Private Function getMailInfo() As String
    Dim i As Long
    
    For i = 1 To mailInfo.Count
        getMailInfo = getMailInfo & IIf(i > 1, vbCrLf, vbNullString) & mailInfo(i)
    Next
    
    Set mailInfo = New Collection
    
End Function

'�u���C���v�V�[�g�u�l���f�[�^��UL�v�{�^����������A���s
'���s�O��Ƃ��āA�ΏۃZ�����u���C���v�V�[�g�\�̍��ږ��u�Ώۊ�Ɩ��v�f�[�^��I�����Ă��鎖������
'�t�@�C�����w�肵�Ă��炢�A�t�@�C���̑Ώۊ�Ƃ�I�����Ă��������AUpLoad���������s
Public Sub upPsDataOnly()
    Dim i As Long
    Dim corpName As String '�Ώۊ�Ɩ�
    Dim csvPath As String  '�Ώ�CS���t�@�C���t���p�X
    Dim ret As Long
    
    '�����W���[���uMain�v�̃v���V�[�W���upreCheck�v�́A3�̑O���������s
    '�@���s�Ҏ����̃f�[�^�L���`�F�b�N�A�A�ΏۃV�[�g(�u���C���v�u�A�J�E���g�v�u�ߋ����O�v�u���[���A�J�E���g�v�V�[�g)�V�[�g�ی�`�F�b�N�A
    '�B�ΏۃV�[�g(�u���s���O�v�u�ߋ����O�v)�̃t�B���^�[����
    If Not preCheck Then
        Exit Sub
    End If
    
    cancelFlg = False
      
    Set opeLog = New Collection
    Set mailInfo = New Collection
    
    corpName = Cells(Selection.row, 2)  '�Ώۊ�Ɩ����擾(�Ώۊ�Ɩ��̃Z����I�����Ă���O��)
    
    If corpName = vbNullString Then Exit Sub  '�ΏۃZ�����A�u���C���v�V�[�g�̍��ږ��u�Ώۊ�Ɩ��v�f�[�^��I�����Ă��Ȃ��ꍇ�́A�����I��
    If MsgBox("�Ώۊ�Ƃ�" & corpName & "�ł�낵���ł����H", vbYesNo) <> vbYes Then Exit Sub
    
    MsgBox "�l���A�b�v���[�h����CSV�t�@�C����I�����Ă��������B"
    '���W���[���uFileOpe�v�́ugetFilePathByDialog�v�̏��������s
    'CSV�t�@�C�����A���[�U�[�ɑI�����Ă��炤
    csvPath = getFilePathByDialog("*.csv", "CSV�t�@�C��", "�l���A�b�v���[�h����t�@�C����I�����Ă��������B")
    If csvPath = vbNullString Then Exit Sub
    
    ret = MsgBox("�ΏۃT�C�g�̓}�C�i�r�ł����H" & vbCrLf & "���N�i�r�Ȃ�u������(N)�v��I��", vbYesNoCancel)
    
    If ret = vbYes Then
        upMyNaviPsDataOnly corpName, csvPath  '�}�C�i�r��UpLoad������
    ElseIf ret = vbNo Then
        upRikuNaviPsDataOnly corpName, csvPath  '���N�i�r��UpLoad������
    Else
        Exit Sub
    End If
    
    opeLog.Add "i-Web�C���|�[�g���������܂����I"
    
    Dim msg As String
    
    For i = 1 To opeLog.Count
        msg = msg & IIf(msg = vbNullString, vbNullString, vbCrLf) & opeLog(i)
    Next
    
    If Not msg = vbNullString Then
        MsgBox msg, vbInformation
    End If
    
    Set opeLog = Nothing
    
    Unload AlertBox

End Sub

Private Function upRikuNaviPsDataOnly(ByVal tgtCorpName As String, _
                                      ByVal csvPath As String) As Boolean

'###IE���b�p�[���N��
    '##i-Web�p��IE���b�p�[���N��
    Dim iweb As CorpSite
    Set iweb = New CorpSite
    
    On Error GoTo wrapperErr
    '���O�ǉ�
    opeLog.Add "InternetExplore���N����..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr
    iweb.cleanUpTgtSite
       
    '##���N�i�r�p��IE���b�p�[���N��
    Dim rikuNavi As CorpSite
    Dim rikuNaviFlg As Boolean
    Set rikuNavi = New CorpSite

    If rikuNavi Is Nothing Then GoTo wrapperErr
    If rikuNavi.setCorp(tgtCorpName, "���N�i�r") Then
        rikuNaviFlg = True
        rikuNavi.cleanUpTgtSite
    End If
    
   '���O�o��
    opeLog.Add "�N�������B"
    
    On Error GoTo 0
    
    '�p�X�o�^
    Dim rkNavPsDataFilePath As String
    Dim rkNavPsFlg As Boolean
    
    opeLog.Add "�y�蓮�z���N�i�r�̓��O�C�����܂���B�o�^���ꂽ�f�[�^���g���܂��B"
            
    rkNavPsDataFilePath = csvPath
    rkNavPsFlg = True

    On Error GoTo ulErr
'###���N�i�r�̏���i-Web�փA�b�v���[�h
    If Not rkNavPsDataFilePath = vbNullString And rkNavPsFlg Then

        AlertBox.Label1 = "���N�i�r��CSV����AiWeb�֌l�����A�b�v���[�h���Ă��܂��B"
        opeLog.Add "���N�i�r��CSV����AiWeb�֌l�����A�b�v���[�h��.."
        mailInfo.Add "(���N�i�r)"

        If Not ULPersonalData(iweb, rkNavPsDataFilePath, rikuNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "�A�b�v���[�h����"
    Else
        opeLog.Add "(���N�i�r)" & vbCrLf & "�Ώێ҂����܂���ł����B"
    End If
    On Error GoTo 0
    
normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not rikuNavi Is Nothing Then rikuNavi.quitAll
    
Exit Function

wrapperErr:
    upRikuNaviPsDataOnly = False
    GoTo normalFin
   
ulErr:
    upRikuNaviPsDataOnly = False
    GoTo normalFin

End Function

'������2��
'�@�ΏۃZ�����I�����Ă���u���C���v�V�[�g�̍��ږ��u�Ώۊ�Ɩ��v�f�[�^�A�A�I�����Ă������CSV�t�@�C���t���p�X
Private Function upMyNaviPsDataOnly(ByVal tgtCorpName As String, _
                                    ByVal csvPath As String) As Boolean

'###IE���b�p�[���N��
    '##i-Web�p��IE���b�p�[���N��
    Dim iweb As CorpSite
    Set iweb = New CorpSite  'Index�Ɂu1�v�������āAIE�I�u�W�F�N�g���擾
    
    On Error GoTo wrapperErr
    '���O�ǉ�
    opeLog.Add "InternetExplore���N����..."

    If iweb Is Nothing Then GoTo wrapperErr
    If Not iweb.setCorp(tgtCorpName, "i-Web") Then GoTo wrapperErr 'iweb�擾
    iweb.cleanUpTgtSite
       
    '##�}�C�i�r�p��IE���b�p�[���N��
    Dim myNavi As CorpSite
    Dim myNaviFlg As Boolean
    Set myNavi = New CorpSite

    If myNavi Is Nothing Then GoTo wrapperErr
    If myNavi.setCorp(tgtCorpName, "�}�C�i�r") Then  '�}�C�i�r�擾
        myNaviFlg = True
        myNavi.cleanUpTgtSite
    End If
    
   '���O�o��
    opeLog.Add "�N�������B"
    
    On Error GoTo 0
    
    '�p�X�o�^
    Dim myNavPsDataFilePath As String
    Dim myNavPsFlg As Boolean
    
    opeLog.Add "�y�蓮�z�}�C�i�r�̓��O�C�����܂���B�o�^���ꂽ�f�[�^���g���܂��B"
            
    myNavPsDataFilePath = csvPath
    myNavPsFlg = True

    On Error GoTo ulErr
'###�}�C�i�r�̏���i-Web�փA�b�v���[�h
    If Not myNavPsDataFilePath = vbNullString And myNavPsFlg Then

        AlertBox.Label1 = "�}�C�i�r��CSV����AiWeb�֌l�����A�b�v���[�h���Ă��܂��B"
        opeLog.Add "�}�C�i�r��CSV����AiWeb�֌l�����A�b�v���[�h��.."

        If Not ULPersonalData(iweb, myNavPsDataFilePath, myNavi.NavPDLayNo) Then GoTo ulErr

        opeLog.Add "�A�b�v���[�h����"
    Else
        opeLog.Add "���l���C���|�[�g��" & vbCrLf & "(�}�C�i�r)" & vbCrLf & "�Ώێ҂����܂���ł����B"
    End If
    On Error GoTo 0
    
normalFin:
    If Not iweb Is Nothing Then iweb.quitAll
    If Not myNavi Is Nothing Then myNavi.quitAll
    
Exit Function

wrapperErr:
    upMyNaviPsDataOnly = False
    GoTo normalFin
   
ulErr:
    upMyNaviPsDataOnly = False
    GoTo normalFin

End Function

