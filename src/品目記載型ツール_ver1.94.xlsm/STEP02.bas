Attribute VB_Name = "STEP02"
Option Explicit

Sub STEP02���W���[��()

    Call ���䕶���(Pub����V�[�g��)
    
End Sub

Private Sub ���䕶���Test()
    ���䕶��� WSNAME_WARIKOMI
End Sub

Public Sub ���䕶���(ByVal ����V�[�g�� As String)

    '�ċA�����t�v���V�[�W���i�����ӂ��I�j

    Dim MaxRow As Long
    
    MaxRow = ThisWorkbook.Sheets(����V�[�g��).Cells(Rows.Count, 10).End(xlUp).Row
    
    Dim �E�B���h�E As String, �A�N�V���� As String, �^�O As String, �^�� As String, �N���b�N�Ώ� As String, �I�v�V����1 As String, �I�v�V����2 As String
    Dim �e�L�X�g As String

    Dim i As Long
    Dim �Ώۃt�@�C���p�X As String
'    Dim �Ώۃt�@�C���� As String
    Dim CSVFile As Workbook
    For i = 20 To MaxRow
        
        DoEvents
'        If fStop = True Then GoTo HandleError
        
        '----------------------------------------------------
        '���R�[�h�I�𔻒�
        '----------------------------------------------------
        If ThisWorkbook.Sheets(����V�[�g��).Cells(i, 7).Value <> "��" Then
            If ThisWorkbook.Sheets(����V�[�g��).Cells(i, 8).Value <> "��" Then GoTo continue
        End If
        
        '----------------------------------------------------
        '�f�[�^�擾
        '----------------------------------------------------
        Pub����ԍ� = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 9).Value
        
        �E�B���h�E = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 10).Value
        �A�N�V���� = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 11).Value
        �^�O = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 12).Value
        �^�� = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 13).Value
        �N���b�N�Ώ� = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 14).Value
        �I�v�V����1 = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 15).Value
        �I�v�V����2 = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 16).Value
        
        �e�L�X�g = ThisWorkbook.Sheets(����V�[�g��).Cells(i, 20).Value
        
        '----------------------------------------------------
        '�A�N�V��������i��~�j
        '----------------------------------------------------
        If �A�N�V���� = "��~" Then MessageBox 0, "��~�A�N�V����", "�I�[�g�p�C���b�g��~", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION: End
        
        '----------------------------------------------------
        '�A�N�V��������i�f�[�^�擾�j
        '----------------------------------------------------
        If �A�N�V���� = "�f�[�^�擾" Then
            Pub�Ώۃt�@�C���p�X = GetElementByID(�E�B���h�E, �N���b�N�Ώ�)
'            Pub�Ώۃt�@�C���p�X = "C:\Users\21501173\Desktop\NESIC�l\PSI�M�C �l�b�g���[�N�C���t���\�z_�yWi-Fi�\�z�E�ێ�^�p�z.csv"
            If Pub�Ώۃt�@�C���p�X <> "" Then
                Set CSVFile = Workbooks.Open(Pub�Ώۃt�@�C���p�X)
                Pub���ϓo�^���� = CSVFile.Worksheets(1).Range("K3").Value
                Pub���� = Pub���ϓo�^����
                Pub�c�Ǝ҃R�[�h = CSVFile.Worksheets(1).Range("D3").Value
'                Pub��C�҃R�[�h = CSVFile.Worksheets(1).Range("E3").Value
                Pub��C�҃R�[�h = CSVFile.Worksheets(1).Range("P3").Value
                Pub�H��FROM = CSVFile.Worksheets(1).Range("M3").Value
                Pub�H��TO = CSVFile.Worksheets(1).Range("N3").Value
                Pub���ϑO����� = TextFileData(GetTextFileFullPath(Pub�Ώۃt�@�C���p�X))
                
                SetKengyoHo CSVFile.Worksheets(1)
                SetChotatsukubun CSVFile.Worksheets(1)
                SetNextPage
                
                CSVFile.Close False
                GoTo continue
            End If
        End If
        '----------------------------------------------------
        '�A�N�V��������i�ϐ����́j
        '----------------------------------------------------
        If �A�N�V���� = "�ϐ�����" Then
            
            �e�L�X�g = IIf(�e�L�X�g = "������", Pub���ϓo�^����, �e�L�X�g)
            �e�L�X�g = IIf(�e�L�X�g = "���c�Ǝ҃R�[�h", Pub�c�Ǝ҃R�[�h, �e�L�X�g)
            If �e�L�X�g = "����C�҃R�[�h" Then
                If IsEmpty(Pub�H��FROM) = True Or IsNull(Pub�H��FROM) = True Or Pub�H��FROM = "" Then
                    �e�L�X�g = vbNullString
                Else
                    �e�L�X�g = IIf(�e�L�X�g = "����C�҃R�[�h", Pub��C�҃R�[�h, �e�L�X�g)
                End If
            End If
            �e�L�X�g = IIf(�e�L�X�g = "���H��FROM", Format(Pub�H��FROM, "ddddd"), �e�L�X�g)
            �e�L�X�g = IIf(�e�L�X�g = "���H��TO", Format(Pub�H��TO, "ddddd"), �e�L�X�g)
            �e�L�X�g = IIf(�e�L�X�g = "�����ϑO�����", Pub���ϑO�����, �e�L�X�g)
            If �e�L�X�g = "" Then GoTo continue

        End If
        
        '----------------------------------------------------
        '�A�N�V��������
        '----------------------------------------------------
        If �A�N�V���� = "�V�KIE" Then
            Call �V�KIE�J��(�E�B���h�E)
            IE�O�� �E�B���h�E
            GoTo continue
        End If
        If �A�N�V���� = "����IE" Then
            Call ����IE�J��(�E�B���h�E)
            IE�O�� �E�B���h�E
            GoTo continue
        End If
        If �A�N�V���� = "LUCAS���O�C��������" Then LucasLoginForm.Show: GoTo continue
        If �A�N�V���� = "LUCAS���O�C��" Then
            Call ieInTextBoxTagInputTypeText(�E�B���h�E, �A�N�V����, "input", "USER_INPUT", �I�v�V����1, �I�v�V����2, PubClsLucasAuth.LucasID)
            Call ieInTextBoxTagInputTypeText(�E�B���h�E, �A�N�V����, "input", "PASSWORD", �I�v�V����1, �I�v�V����2, PubClsLucasAuth.LucasPW)
            Call ieClickSubmitButtonTagInputTypeSubmit(�E�B���h�E, �A�N�V����, "input", "���O�C��", �I�v�V����1, �I�v�V����2): GoTo continue
        End If
        
        
        
        If �A�N�V���� = "���O�C���m�F" Then Call ieCheckSSISLogin(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
        
        If �A�N�V���� = "���Ϗo��" Then Call ��O_���Ϗo��.��O_���Ϗo�̓��W���[��: GoTo continue
        
        If �A�N�V���� = "�ʒm�o�[" Then Call ieDownloadFileNbOrDlg(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
        
        '        If �A�N�V���� = "�l���o" Then
        '            If �^�� = "hidden" Then Call ieExValueTagInputTypeHidden(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
        '        End If
        
        '----------------------------------------------------
        '�A�N�V��������i�����j
        '----------------------------------------------------
        If �A�N�V���� = "����" Then
            Call ���䕶���(�N���b�N�Ώ�) '�ċA����
            IE�O�� �E�B���h�E
        End If
         
        '----------------------------------------------------
        '�^�O����^�`������
        '----------------------------------------------------
        If �^�O = "a" Then
            If i = MaxRow Then
                Pub���ϔԍ� = ���ϔԍ��擾
            End If
            Call ieClickLinkTagAhrefTypeNone(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2): GoTo continue
            IE�O�� �E�B���h�E
        End If
        
        If �^�O = "input" Then
            If �^�� = "button" Then Call ieClickButtonTagInputTypeButton(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2): GoTo continue
            If �^�� = "checkbox" Then Call ieClickCheckBoxTagInputTypeCheckBox(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2): GoTo continue
            If �^�� = "text" Then Call ieInTextBoxTagInputTypeText(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
            If �^�� = "file" Then Call ieClickButtonTagInputTypeFile(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
            If �^�� = "submit" Then Call ieClickSubmitButtonTagInputTypeSubmit(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2): GoTo continue
            If �^�� = "radio" Then Call ieClickRadioButtonTagInputTypeRadio(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2): GoTo continue
            If �^�� = "password" Then Call ieInPasswordBoxTagInputTypePassword(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
            IE�O�� �E�B���h�E
        End If
        
        If �^�O = "textarea" Then
            If �^�� = "text" Then Call ieInTextTagTextAreaTypeText(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
            IE�O�� �E�B���h�E
        End If
        

        If �^�O = "select" Then
            Call ieClickSelectBoxTagSelect(�E�B���h�E, �A�N�V����, �^�O, �N���b�N�Ώ�, �I�v�V����1, �I�v�V����2, �e�L�X�g): GoTo continue
            IE�O�� �E�B���h�E
        End If
continue:

    Next
    
    Exit Sub
    
HandleError:

    MessageBox 0, "�I�[�g�p�C���b�g�͒�~���܂����B", "�m�F", MB_OK Or MB_TOPMOST Or MB_EXCLAMATION
    
    End
End Sub

