Attribute VB_Name = "M02_CSV"
Option Explicit

Public Pub�Ώۃt�@�C���p�X As String
Public Pub���� As String
Public Pub���ϔԍ� As String
Public Pub���Ɩ@ As String
Public PubHas���Ɩ@ As Boolean
Public Pub�����@ As String
Public PubHas�����@ As Boolean

Public Pub���t As Variant
Public Pub�Ј��ԍ�  As Variant

'CSV�t�@�C���쐬
Public Sub CSV�t�@�C���쐬(ByVal Target As Range)
    Dim �^�C�g�� As Variant
    Dim ��ʃt���O As String
    Dim ���ς��茏�� As String
    Dim �p�^�[�� As String
    Dim �J�n�� As Variant
    Dim �I���� As Variant
    
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHdl
    
    �^�C�g�� = Target.Offset(, 2).Value
    ��ʃt���O = Target.Offset(, 3).Value
    Dim ������Z�t���O As Boolean
    If Target.Offset(0, 3).Value Like "*������Z*" Then
        ������Z�t���O = True
    Else
        ������Z�t���O = False
    End If
    �p�^�[�� = Target.Offset(, 5).Value
    Dim �r�W�l�XIT As String
    �r�W�l�XIT = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE).Range("AU3").Value
    '    ���ς��茏�� = Target.Offset(, 4).Value
    
    If �K�{�f�[�^�`�F�b�N(Target, �p�^�[��) = False Then
        Exit Sub
    End If
    
    
    If IS_TEST = False Then
        ���䕶��� WSNAME_LOGIN
    End If
    '
    '�ۑ���`�F�b�N
    Dim SavePath As String
    SavePath = �ݒ�擾("���i�ڋL���^", "csv�t�@�C���ۑ���t�H���_")
    If HasTargetFolder(SavePath) = False Then
        MsgBox "csv�t�@�C���ۑ���t�H���_���u�ݒ�v�V�[�g�Ɏw�肵�Ă�������" _
            , vbExclamation
        Exit Sub
    End If

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    sh.Rows(3).Clear
    Dim vRow As Long
    vRow = sh.Cells.SpecialCells(xlCellTypeLastCell).Row
    If vRow < 6 Then
        vRow = 6
    End If
    sh.Rows("6:" & sh.Rows.Count).Clear
    '    sh.Range("A5").CurrentRegion.Offset(1).Clear
    Dim ���ϑO����� As String
    ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value = ""

    ���W���[���]�L �p�^�[��

    Dim i As Long, j As Long
    i = 6
    Do
        DoEvents
        If sh.Cells(i, 1).Value = "" Then
            Exit Do
        End If
        �i�ڏ��]�L sh.Cells(i, 1).Value, i, ������Z�t���O
        i = i + 1
    Loop

    Dim �ϊ��Ώ� As Range
    Dim LastRow As Long
    With Sheet0
        LastRow = .Cells(.Rows.Count, 15).End(xlUp).Row
        If LastRow < 6 Then
            LastRow = 6
        End If
        Set �ϊ��Ώ� = .Range(.Cells(6, 15), .Cells(LastRow, 19))
    End With
    
    Dim HasKengyoHo As Boolean
    Dim HasShitaukeHo As Boolean
    Dim EndRow As Long
    Dim vBook As Workbook
    '    Set vBook = Workbooks.Add
    '    ThisWorkbook.Worksheets(WSNAME_FORMAT).Cells.Copy _
    '        vBook.Worksheets(1).Range("A1")
    Dim ���σV�[�g As Worksheet
    Dim vMsg As String
    Set ���σV�[�g = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    For i = 1 To �ϊ��Ώ�.Rows.Count
        DoEvents
        HasKengyoHo = False
        Pub���Ɩ@ = vbNullString
        Pub�����@ = vbNullString
        ���ς��茏�� = Target.Offset(, 4).Value
        ���ϑO����� = ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value

        If Target.Offset(, 6).Value = "" Then
            ���ς��茏�� = Replace(���ς��茏��, "����", �ϊ��Ώ�(i, 1))
            ���ϑO����� = Replace(���ϑO�����, "����", �ϊ��Ώ�(i, 1))
        Else
            ���ς��茏�� = Replace(���ς��茏��, "����", Target.Offset(, 6).Value)
            ���ϑO����� = Replace(���ϑO�����, "����", Target.Offset(, 6).Value)
        End If
        If Target.Offset(, 7).Value = "" Then
            ���ς��茏�� = Replace(���ς��茏��, "����", �ϊ��Ώ�(i, 2))
            ���ϑO����� = Replace(���ϑO�����, "����", �ϊ��Ώ�(i, 2))
        Else
            ���ς��茏�� = Replace(���ς��茏��, "����", Target.Offset(, 7).Value)
            ���ϑO����� = Replace(���ϑO�����, "����", Target.Offset(, 7).Value)
        End If
        If Target.Offset(, 8).Value = "" Then
            ���ς��茏�� = Replace(���ς��茏��, "����", �ϊ��Ώ�(i, 3))
            ���ϑO����� = Replace(���ϑO�����, "����", �ϊ��Ώ�(i, 3))
        Else
            ���ς��茏�� = Replace(���ς��茏��, "����", Target.Offset(, 8).Value)
            ���ϑO����� = Replace(���ϑO�����, "����", Target.Offset(, 8).Value)
        End If
        If Target.Offset(, 9).Value = "" Then
            ���ς��茏�� = Replace(���ς��茏��, "����", �ϊ��Ώ�(i, 4))
            ���ϑO����� = Replace(���ϑO�����, "����", �ϊ��Ώ�(i, 4))
        Else
            ���ς��茏�� = Replace(���ς��茏��, "����", Target.Offset(, 9).Value)
            ���ϑO����� = Replace(���ϑO�����, "����", Target.Offset(, 9).Value)
        End If
        If Target.Offset(, 10).Value = "" Then
            ���ς��茏�� = Replace(���ς��茏��, "����", �ϊ��Ώ�(i, 5))
            ���ϑO����� = Replace(���ϑO�����, "����", �ϊ��Ώ�(i, 5))
        Else
            ���ς��茏�� = Replace(���ς��茏��, "����", Target.Offset(, 10).Value)
            ���ϑO����� = Replace(���ϑO�����, "����", Target.Offset(, 10).Value)
        End If
            
        Set vBook = Workbooks.Add
        ThisWorkbook.Worksheets(WSNAME_FORMAT).Cells.Copy _
            vBook.Worksheets(1).Range("A1")
        With vBook.Worksheets(1)
            EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            For j = 5 To .UsedRange.Rows.Count
                If Len(.Cells(j, 1).Value) = 0 Then
                    EndRow = j - 1
                    Exit For
                End If
            Next
            .Rows(EndRow + 1 & ":" & .Rows.Count).Delete
        End With
        For j = 5 To EndRow
            If vBook.Worksheets(1).Cells(j, 26).Value = 1 Then
                HasKengyoHo = True
                Pub���Ɩ@ = "��"
                Exit For
            End If
        Next
        For j = 5 To EndRow
            If vBook.Worksheets(1).Cells(j, 27).Value = 1 Then
                HasShitaukeHo = True
                Pub�����@ = "��"
                Exit For
            End If
        Next
        If HasKengyoHo = True Then
            '            MsgBox "�u" & ���ς��茏�� & "�v�̍�ƊJ�n������͂��Ă�������", vbInformation
            vMsg = "���Ɩ@�Ώۃf�[�^������܂�" _
                & vbCrLf & "�u" & ���ς��茏�� & "�v�̍�ƊJ�n������͂��Ă�������"
            MessageBox 0, vMsg, "�m�F", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            C01_UF01.Show
            �J�n�� = Pub���t
            '        �J�n�� = InputBox("��ƊJ�n��")
            '            MsgBox "�u" & ���ς��茏�� & "�v�̍�ƏI��������͂��Ă�������", vbInformation
            vMsg = "�u" & ���ς��茏�� & "�v�̍�ƏI��������͂��Ă�������"
            MessageBox 0, vMsg, "�m�F", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            C01_UF01.Show
            �I���� = Pub���t
        End If
        With vBook.Worksheets(1)
            '            EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
            .Range("A3").Value = 1
            .Range("B3").Value = 2
            If Target.Offset(0, 3).Value Like "*������Z*" Then
                .Range("C3").Value = 4
            Else
                .Range("C3").Value = 2
            End If
            .Range("E3").Value = Pub�Ј��ԍ� ' Func�Ј����擾
            .Range("F3").NumberFormat = "YYYY/MM/DD"
            .Range("F3").Value = Format(DateAdd("M", ���σV�[�g.Range("Y3").Value, Date), "YYYY/MM/DD")
            .Range("G3").Value = ���σV�[�g.Range("Y3").Value
            .Range("H3").Value = ���σV�[�g.Range("AF3").Value
            .Range("I3").Value = ���σV�[�g.Range("AK3").Value
            .Range("J3").Value = ���σV�[�g.Range("AP3").Value
            .Range("K3").Value = ���ς��茏��
            .Range("L3").Value = vbNullString   '���ϑO�����
            .Range("M3").NumberFormat = "YYYY/MM/DD"
            .Range("M3").Value = Format(�J�n��, "YYYY/MM/DD")
            .Range("N3").NumberFormat = "YYYY/MM/DD"
            .Range("N3").Value = Format(�I����, "YYYY/MM/DD")
            .Range("O3").Value = �r�W�l�XIT
        End With

        'CSV�o��
        vBook.SaveAs Filename:=SavePath & "\" & ���ς��茏�� & ".csv", _
            FileFormat:=xlCSV
        
        vBook.Close False

        Open SavePath & "\" & ���ς��茏�� & ".txt" For Output As #1

        Print #1, ���ϑO�����
 
        Close #1
    Next
    
    Dim �Ώۃt�@�C���� As Long
    �Ώۃt�@�C���� = GetFileCount(SavePath, "csv")
    
    For i = 1 To �Ώۃt�@�C����
        If IS_TEST = False Then
            ���䕶��� WSNAME_CSVUP
            �Ώۃt�@�C���ړ� Pub�Ώۃt�@�C���p�X
            
        End If
        ���s���ʋL�^
        '        ���ϔԍ��t�@�C���ۑ� SavePath & "\old", Pub����, Pub���ϔԍ�
    Next
    
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Worksheets(WSNAME_LOG)
    
    Dim MsgRange1 As Range
    Dim MsgRange2 As Range
    
    Set MsgRange1 = wsResult.Columns("K").Find("���Ɩ@�����鎞�̕���")
    Set MsgRange2 = wsResult.Columns("L").Find("�����@�����鎞�̕���")
    
    If PubHas���Ɩ@ Then
        If MsgRange1 Is Nothing Then
        Else
            MessageBox 0, MsgRange1.Offset(1).Value _
                , "�m�F", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
        End If
    End If
    If PubHas�����@ Then
        If MsgRange2 Is Nothing Then
        Else
            
            MessageBox 0, MsgRange2.Offset(1).Value _
                , "�m�F", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
        End If
    End If
    ThisWorkbook.Worksheets(WSNAME_FORMAT).Range("L3").Value = vbNullString
    MsgBox "�������I�����܂���", vbInformation
    
    Exit Sub
ErrHdl:
    MsgBox Err.Description, vbExclamation
End Sub

'�K�{�f�[�^�̓��̓`�F�b�N
Private Function �K�{�f�[�^�`�F�b�N(ByVal Target As Range _
    , ByVal �p�^�[�� As Variant) As Boolean
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    Dim temp As String
    If Target.Offset(, 4) = "" Then
        temp = temp & "�u�����v"
    End If
    If sh.Range("Y3").Value = "" Then
        temp = temp & "�u���ώ��[���v"
    End If
    If sh.Range("AF3").Value = "" Then
        temp = temp & "�u���ϔ[���P�ʁv"
    End If
    If sh.Range("AU3").Value = "" Then
        temp = temp & "�u�r�W�l�XIT�v"
    End If
    Dim vData As Variant
    vData = ���W���[�����e�擾("�c�Ǝ�S���҃R�[�h", "�ϐ���_����", "A")
    If Trim(vData) = "" Then
        temp = temp & "�u�c�Ǝ�S���҃R�[�h�v"
    End If
    vData = ���W���[�����e�擾("�c�Ǝ�S���҃R�[�h", "�ϐ���_����", "A")

    If Len(temp) > 0 Then
        MsgBox "�K�{���ڂ�" & temp & "�������͂ł�" _
            & "�m�F���Ă�������", vbExclamation
        �K�{�f�[�^�`�F�b�N = False
    Else
        �K�{�f�[�^�`�F�b�N = True
    End If
End Function

'Log�̋L�^
Private Sub ���s���ʋL�^()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_LOG)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 5).End(xlUp).Offset(1).Row
        .Cells(LastRow, 5).Value = Format(Now, "YYYY/MM/DD HH:NN")
        .Cells(LastRow, 6).Value = Pub����
        .Cells(LastRow, 7).Value = Pub���ϔԍ�
        .Cells(LastRow, 8).Value = Pub���Ɩ@
        .Cells(LastRow, 9).Value = Pub�����@
        If Pub���Ɩ@ = "��" Then
            PubHas���Ɩ@ = True
        End If
        If Pub�����@ = "��" Then
            PubHas�����@ = True
        End If
    End With
End Sub

'���W���[�����e�̓]�L
Private Sub ���W���[���]�L(ByVal �p�^�[�� As String)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE)
    Dim LastRow As Long
    
    With sh
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 15 To LastRow
        If sh.Cells(i, 3).Value <> "" Then
        ���W���[�����e���� sh.Cells(i, 1).Value, "�ϐ���_����", �p�^�[��
        End If
    Next
End Sub

'�i�ڏ��̓]�L
Private Sub �i�ڏ��]�L(ByVal �i�� As String _
    , ByVal �Ώۍs As Long _
    , ByVal ������Z�t���O As Boolean)
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_HINMOKU)
    For i = 1 To sh.UsedRange.Rows.Count
        DoEvents
        If sh.Cells(i, 1).Value = �i�� Then
            �i�ڊǗ��\������ "�i��", �Ώۍs, i
            �i�ڊǗ��\������ "�T�[�r�X��", �Ώۍs, i
            �i�ڊǗ��\������ "�Œ莑�Y�i�ΏہE��Ώہj", �Ώۍs, i
            �i�ڊǗ��\������ "�����_��X�V�i�ΏہE��Ώہj", �Ώۍs, i
            �i�ڊǗ��\������ "�{��R�[�h", �Ώۍs, i
            �i�ڊǗ��\������ "�̔��P��", �Ώۍs, i
            �i�ڊǗ��\������ "���ϒP��_1�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���ϒP��_2�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���ϒP��_3�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���ϒP��_4�N�ڈȍ~", �Ώۍs, i
            �i�ڊǗ��\������ "�񋟗����敪", �Ώۍs, i
            �i�ڊǗ��\������ "���B���@", �Ώۍs, i
            �i�ڊǗ��\������ "���B�P��", �Ώۍs, i
            �i�ڊǗ��\������ "���B�P��_1�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���B�P��_2�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���B�P��_3�N��", �Ώۍs, i
            �i�ڊǗ��\������ "���B�P��_4�N�ڈȍ~", �Ώۍs, i
            �i�ڊǗ��\������ "���B�����敪", �Ώۍs, i
            �i�ڊǗ��\������ "�Г����B�敪", �Ώۍs, i
            �i�ڊǗ��\������ "��]�Ǝ�", �Ώۍs, i
            �i�ڊǗ��\������ "�����@", �Ώۍs, i
            �i�ڊǗ��\������ "���Ɩ@", �Ώۍs, i
            �i�ڊǗ��\������ "TSC�ێ�L", �Ώۍs, i
            �i�ڊǗ��\������ "���l", �Ώۍs, i
            If ������Z�t���O Then
            �i�ڊǗ��\������ "��������i������Z�̂ݎg�p�j", �Ώۍs, i
            Else
            �i�ڊǗ��\������ "��������i������Z�̂ݎg�p�j", �Ώۍs, i, True
            End If
            Exit Sub
        End If
    Next
End Sub

'�i�ڊǗ��\������
Private Sub �i�ڊǗ��\������(ByVal �Ώ� As String _
    , ByVal �Ώۍs As Long, ByVal �i�ڍs As Long _
    , Optional ByVal flg As Boolean = False)
    Dim i As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_HINMOKU)
    
    For i = 4 To sh.UsedRange.Rows.Count
        DoEvents
        If sh.Cells(i, 1).Value = �Ώ� Then
            
            If flg Then
                ThisWorkbook.Worksheets(sh.Cells(i, 6).Value).Cells(�Ώۍs, sh.Cells(i, 8).Value).Value _
                    = ""
            Else
                With ThisWorkbook.Worksheets(sh.Cells(i, 6).Value).Cells(�Ώۍs, sh.Cells(i, 8).Value)
                    .NumberFormat = "@"
                    .Value = CStr(ThisWorkbook.Worksheets(WSNAME_HINMOKU).Cells(�i�ڍs, sh.Cells(i, 5).Value).Value)
                End With
            End If
            Exit Sub
        End If
    Next
End Sub

Private Sub ���W���[�����e����Test()
    ���W���[�����e���� "�i��2", "�ϐ���_����", "A"
End Sub
'���W���[���̓��e�����
Private Sub ���W���[�����e����(ByVal �Ώ� As String _
    , ByVal ��� As String, Optional �p�^�[�� As String)
    Dim �Ώ۔͈� As Range
    Set �Ώ۔͈� = �Ώ۔͈͎擾(���)
    Dim vData As Variant
    vData = �Ώ۔͈�.Value
    
    Dim �Ώۍs As Long
    Dim i As Long
    For i = LBound(vData) To UBound(vData)
        DoEvents
        If vData(i, 1) = �Ώ� Then
            �Ώۍs = i
            Exit For
        End If
    Next
    If �Ώۍs = 0 Then Exit Sub
    
    If InStr(���, "����") > 0 Then
        Dim ��Z�� As Range
        Set ��Z�� = �p�^�[������(�p�^�[��)
        With ThisWorkbook.Worksheets(vData(�Ώۍs, 7)).Cells(vData(�Ώۍs, 8), vData(�Ώۍs, 9))
            .NumberFormat = "@"
            .Value = CStr(��Z��.Cells(vData(�Ώۍs, 5), vData(�Ώۍs, 6)).Value)
        End With
    Else
        With ThisWorkbook.Worksheets(vData(�Ώۍs, 5)).Cells(vData(�Ώۍs, 6), vData(�Ώۍs, 7))
            .NumberFormat = "@"
            .Value = CStr(.Worksheets(vData(�Ώۍs, 3)).Range(vData(�Ώۍs, 4)).Value)
        End With
    End If
End Sub

Private Sub ���W���[�����e�擾Test()
    Debug.Print ���W���[�����e�擾("�c�Ǝ�S���҃R�[�h", "�ϐ���_����", "A")
End Sub
Private Function ���W���[�����e�擾(ByVal �Ώ� As String _
    , ByVal ��� As String, Optional �p�^�[�� As String) As Variant
    Dim �Ώ۔͈� As Range
    Set �Ώ۔͈� = �Ώ۔͈͎擾(���)
    Dim vData As Variant
    vData = �Ώ۔͈�.Value
    
    Dim �Ώۍs As Long
    Dim i As Long
    For i = LBound(vData) To UBound(vData)
        DoEvents
        If vData(i, 1) = �Ώ� Then
            �Ώۍs = i
            Exit For
        End If
    Next
    If �Ώۍs = 0 Then Exit Function
    
    If InStr(���, "����") > 0 Then
        Dim ��Z�� As Range
        Set ��Z�� = �p�^�[������(�p�^�[��)
        With ThisWorkbook
            ���W���[�����e�擾 = ��Z��.Cells(vData(�Ώۍs, 5), vData(�Ώۍs, 6)).Value
        End With
    Else
        With ThisWorkbook
            ���W���[�����e�擾 = .Worksheets(vData(�Ώۍs, 3)).Range(vData(�Ώۍs, 4)).Value
        End With
    End If
End Function

'�u�ϐ��ꗗ_���W���[���v�V�[�g�̑Ώ۔͈͂̎擾
Private Function �Ώ۔͈͎擾(ByVal ��� As String) As Range
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE)
    
    Dim i As Long
    For i = 1 To sh.UsedRange.Rows.Count
        If sh.Cells(i, 1).Value = ��� Then
            Set �Ώ۔͈͎擾 = sh.Cells(i, 1).CurrentRegion
            Exit Function
        End If
    Next
    
End Function

Private Sub �p�^�[������Test()
    Debug.Print �p�^�[������("A").Address
End Sub
'�u�i�ڋL���^�i���W���[���j�v�V�[�g����p�^�[��������
Private Function �p�^�[������(ByVal �p�^�[�� As String) As Range
    Dim WS�����Ώ� As Worksheet
    Set WS�����Ώ� = ThisWorkbook.Worksheets(WSNAME_HINMOKU_MODULE)
    
    Dim �Ώۗ� As Long
    �Ώۗ� = ThisWorkbook.Worksheets(WSNAME_VAL_MODULE).Range("D4").Value

    Dim �ŏI�s As Long
    With WS�����Ώ�
        �ŏI�s = .Cells(.Rows.Count, �Ώۗ�).End(xlUp).Row
    End With
    
    Dim i As Long
    For i = 1 To �ŏI�s
        If WS�����Ώ�.Cells(i, �Ώۗ�).Value = �p�^�[�� Then
            Set �p�^�[������ = WS�����Ώ�.Cells(i, �Ώۗ�)
            Exit Function
        End If
    Next
End Function

'�u���Ɩ@�v���`�F�b�N����
Public Sub SetKengyoHo(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 26).Value = 1 Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "chk_kengyouhou_" & num Then
                        sh.Cells(j, 8).Value = "��"
                    End If
                Next
            End If
        Next
    End With
End Sub

Private Sub SetFlgTest()
    SetKengyoHo ActiveSheet
    SetShanaiChotatsu ActiveSheet
    SetChotatsukubun ActiveSheet
    SetNextPage
End Sub
'�u�Г����B�v�`�F�b�N
Public Sub SetShanaiChotatsu(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
'        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 28).Value <> "" Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "sel_chotatsu_hoho_" & num Then
                        sh.Cells(j, 15).Value = .Cells(i, 28).Value
                        sh.Cells(j, 8).Value = "��"
                    End If
                Next
            End If
        Next
    End With
End Sub
'�u���B�敪�v�̓���
Public Sub SetChotatsukubun(ByVal CSVSh As Worksheet)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
'        .Range(.Cells(20, 8), .Cells(LastRow, 8)).Value = vbNullString
    End With
    
'    Dim CSVSh As Worksheet
'    Set CSVSh = ThisWorkbook.Worksheets(WSNAME_FORMAT)
    Dim EndRow As Long
    Dim i As Long, j As Long
    Dim num As Long
    With CSVSh
        EndRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 6 To EndRow
            num = num + 1
            If .Cells(i, 28).Value <> "" Then
                For j = 20 To LastRow
                    If sh.Cells(j, 14).Value = "sel_kmo_shanai_choutatsu_" & num Then
                        sh.Cells(j, 15).Value = .Cells(i, 28).Value
                        sh.Cells(j, 8).Value = "��"
                    End If
                Next
            End If
        Next
    End With
End Sub
'���y�[�W�֑J��
Public Sub SetNextPage()
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(WSNAME_WARIKOMI)
    Dim LastRow As Long
    With sh
        LastRow = .Cells(.Rows.Count, 9).End(xlUp).Row
        .Range(.Cells(20, 7), .Cells(LastRow, 7)).Value = vbNullString
    End With
    
    Dim HasData As Boolean
    Dim TargetRow As Long
    Dim i As Long
    
    '�ŏI�f�[�^���擾
    For i = LastRow To 20 Step -1
        If sh.Cells(i, 8).Value = "��" Then
            TargetRow = i
            Exit For
        End If
    Next
    
    '�ŏI�f�[�^���O�̃y�[�W�J�ڃ{�^���̓N���b�N�Ώ�
    For i = 20 To TargetRow
        If sh.Cells(i, 14).Value = "btn_next" Then
            sh.Cells(i, 7).Value = "��"
        End If
    Next
    
End Sub
