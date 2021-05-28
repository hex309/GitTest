Attribute VB_Name = "Sync_MOD"
Option Explicit
'�������n���W���[��
Dim str_AcDBcn As String

Public Sub Run_Douki_KANRIvew()
'��Ac��Ex�Ǘ��\�r���[����
    Call upd_AcExSync2("�Ǘ��\�o�̓r���[", "T_KANRI", "B10", "T_1")

End Sub

Public Sub Run_Search_KANRI()
'���������e��AccessDB��Excel�f�[�^�Ǘ��\�ҏW_�o�^�V�[�g����
    Dim str_WHERE As String
  
    str_WHERE = Get_WHERE("�Ǘ��\�ҏW�o�^", "B2", "B4")
    Call upd_AcExSync2("�Ǘ��\�ҏW�o�^", "T_KANRI", "B10", "T_1", str_WHERE, 1)
    Call Re_Scrl

End Sub

Public Sub Run_Search_Costum_KANRI()
'���������e��AccessDB��Excel�f�[�^�Ǘ��\�ҏW�o�^�V�[�g����
    Call Run_Search_Costumvew("�Ǘ��\�ҏW�o�^")
    With Sheets("�Ǘ��\�ҏW�o�^")
        .Unprotect
        .Range("C:C").Formula = .Range("C:C").Formula
    End With
    Call St_Lock
    Call Re_Scrl
    ThisWorkbook.Save

End Sub

Public Sub Run_Search_KANRIvew(Optional ByVal str_Stn As String = "�Ǘ��\�o�̓r���[", _
                                                     Optional str_rng As String)
'���������e��AccessDB��Excel�f�[�^�Ǘ��\�o�̓r���[�V�[�g����
    Dim str_WHERE As String
    
    Application.ScreenUpdating = False
    str_WHERE = Get_WHERE(str_Stn, "B2", "B4")
    Call upd_AcExSync2(str_Stn, "T_KANRI", "B10", "T_1", str_WHERE, 1)
    Call Re_Scrl

End Sub

Public Sub Run_Search_Costumvew(Optional ByVal str_Stn As String = "�J�X�^���r���[")
'���������e��AccessDB��Excel�f�[�^�Ǘ��\�o�̓r���[�V�[�g����
    Dim str_WHERE As String
    Dim str_Fild As String
    Dim eCol As Long
    Dim i As Long
    
    eCol = ActiveSheet.Range("B7").End(xlToRight).Column
    str_Fild = ""
    For i = 2 To eCol
        str_Fild = str_Fild & ActiveSheet.Cells(7, i).Value & ","
    Next i
    str_Fild = Left(str_Fild, Len(str_Fild) - 1)
    str_WHERE = Get_WHERE(str_Stn, "B2", "B4")
    Call upd_AcExSync1(str_Stn, "T_KANRI", "T_1", "B10", str_Fild, str_WHERE)
    Call Re_Scrl

End Sub
Public Sub Run_Douki_GAIB()
'��Ac��Ex�O���f�[�^����
    Dim str_Fild As String
    str_Fild = Get_SQLFelds("TG_G_ColList")
    Call upd_AcExSync1("�O���f�[�^", "T_GAIBU1", "F_1", "B8", str_Fild)

End Sub

Public Sub Run_Search_GAIB()
'���������e��AccessDB��Excel�f�[�^�Ǘ��\�V�[�g����
    Dim str_WHERE As String
    Dim str_Fild As String
    str_Fild = Get_SQLFelds("TG_G_ColList")
  
    str_WHERE = Get_WHERE("�O���f�[�^", "B1", "B3")
    If str_WHERE = "" Then End
    
    Call upd_AcExSync1("�O���f�[�^", "T_GAIBU1", "F_1", "B8", str_Fild, str_WHERE)
    Call Re_Scrl

End Sub

Public Function upd_AcExSync1(ByVal str_Stn As String, str_Tbl As String, str_Key As String, _
                                                str_vRng As String, _
                                                Optional str_Fild As String = "*", _
                                                Optional str_WHERE As String = "")
'��Access�O���f�[�^�˓���Excel�V�[�g����
'�@�K�p�V�[�g:�J�X�^��
    '���L��upd_AcExSync2(�Ǘ��\�p)�Ƃ̈Ⴂ:�擾�t�B�[���h�w��@�\������
    '(����1:�������݃V�[�g��,����2:�Ǐo���e�[�u����,����3:Null���O�t�B�[���h�Z���A�h���X
    '  ����4:�\�t���Z���A�h���X,����5: �J�����ݒ�V�[�g��,����6: �擾�t�B�[���h��=�ȗ����͑S�t�B�[���h,
    '����7:SQL�ǉ�������=�ȗ�����"")
    Dim L_Ws As Worksheet
    Dim str_LCStn As String
    Dim str_SQL As String
    Dim eRow As Long
    Dim RcCnt As String
'�Ǐo�f�[�^�Z�b�g Access
    Call Opn_AcRs("T_KANRI", str_Key, str_WHERE, str_Fild)
'�Ǐo�f�[�^�Z�b�g�����܂�
'���f�[�^�]�L
    Set L_Ws = Sheets(str_Stn)

    With L_Ws
        .Unprotect
         .Range("11:20000").Delete
        .Range("B11:GZ20000").ClearContents
        .Range(str_vRng).CopyFromRecordset Ac_Rs
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        If eRow < 10 Then End
        .Range("B10").EntireRow.Copy
        .Range(11 & ":" & eRow).PasteSpecial Paste:=xlPasteFormats
        .Range("G:HZ").EntireColumn.AutoFit
        RcCnt = eRow - 9
        If eRow = 10 Then
            .Range("11:11").Delete
            RcCnt = "1"
        End If
        If str_Stn = "�Ǘ��\�ҏW�o�^" Then
            .Range(eRow + 1 & ":1000").Interior.Color = 16777164
            .Shapes("Rc_Cnt").TextFrame2.TextRange.Characters.Text = RcCnt
        End If
    End With
    Call Dis_Ac_Rs
    Call St_Lock
    Exit Function
Era1:
    If Err.Number = -2147467259 Then
        MsgBox "DB�t�@�C���֐ڑ��ł��܂���ł��� " & vbCrLf & _
         "�f�B���N�g���ݒ�Ńp�X���m�F�E�Đݒ肵�Ă�������" & vbCrLf & _
         "OK�������Ɛݒ�y�[�W�ֈړ����܂�", 16
         Call vis_SETDirectSt
        End
    End If
 
End Function

Public Function upd_AcExSync2(ByVal str_Stn As String, str_Tbl As String, _
                                                str_rng As String, str_Key As String, _
                                                Optional str_WHERE As String = "", Optional Flg As Long = 0)
'��Access�e�[�u���f�[�^�˓���EXcel�V�[�g���� (�K�p�V�[�g�F�Ǘ��\�j
    '(����1:�������݃V�[�g��,����2:�Ǐo���e�[�u����,����3:�\�t���Z���A�h���X,����4:Null�����
     ',����5:�����L�[��,�S���t���O =0:�S��/=�P:�����i�� �ȗ�����0)
    Dim L_Ws As Worksheet
    Dim str_SQL As String
    Dim sRow, sCol, eRow As Long
    '���]�L�p�V�[�g�I�u�W�F�N�g�Z�b�g
    Set L_Ws = Sheets(str_Stn)
    '��������Access�ďo�f�[�^�L���`�F�b�N
    If Flg = 1 Then '�S���t���O=1=�����E�i��
        If str_WHERE <> "" Then '�����������w�肠��
            Call Opn_AcRs(str_Tbl, str_Key, str_WHERE) '����pAccess���R�[�h�Z�b�g
            If Ac_Rs.EOF = True Then '�f�[�^�L�����聁�f�[�^���Ȃ�������
                MsgBox "�f�[�^��������܂���ł���", 16
                Call Dis_Ac_Rs
                End
            End If
        Else '�����������w��Ȃ�
            MsgBox "�����E�i���������w�肵�Ă�������", 16
            End
        End If
    ElseIf Flg = 0 Then '�S���t���O=0=�S��
        Call Opn_AcRs(str_Tbl, str_Key)
    End If
    With L_Ws
        .Unprotect
        sRow = .Range(str_rng).Row
        .Range(sRow + 1 & ":90000").Delete
        .Range(str_rng & ":GZ10000").ClearContents
        .Range(str_rng).CopyFromRecordset Ac_Rs
        sCol = .Range(str_rng).Column
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        .Range("B10").EntireRow.Copy
        .Range(sRow + 1 & ":" & eRow).PasteSpecial Paste:=xlPasteFormats
        If eRow = 10 Then
            .Range("B11").EntireRow.Delete
        End If
        .Range(.Cells(sRow, sCol), .Cells(eRow, 10)).Formula = _
        .Range(.Cells(sRow, sCol), .Cells(eRow, 10)).Formula
        .Range(.Cells(Columns.Count, 7), .Cells(Columns.Count, 200)).EntireColumn.AutoFit
        .Range(eRow + 1 & ":1000").Interior.Color = 16777164
    End With
    Call Dis_Ac_Rs
    Call St_Lock

End Function
