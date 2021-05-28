Attribute VB_Name = "OpnRs_MOD"
Option Explicit
'�����R�[�h�Z�b�g�擾�n���W���[��


Public Function Opn_ExlRs(ByVal str_StRng As String, str_Key As String, _
                                        Optional str_WHERE As String = "", _
                                        Optional str_Fild As String = "*", _
                                        Optional Flg As Long = 0)
'��Excel���R�[�h�Z�b�g�I�[�v��
    '(�����P:�V�[�g�������W,�����Q:��L�[(Null���O�t�B�[���h)�A�����R:�ǉ�������=�ȗ���""
    '�A�����S:�t�B�[���h�w��=�ȗ���"*",����5:�w�b�_�[���L�����w��A�P�Ŗ������ȗ���0�ŗL�j
    Set Exl_Cn = New ADODB.Connection
    Set Exl_Rs = New ADODB.Recordset
    Exl_Cn.Provider = "Microsoft.ACE.OLEDB.12.0"
    If Flg = 0 Then
        Exl_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=YES;IMEX=1"
    ElseIf Flg = 1 Then
        Exl_Cn.Properties("Extended Properties") = "Excel 12.0;HDR=NO;IMEX=1"
    End If
    Exl_Cn.Open ThisWorkbook.FullName
    str_SQL = ""
    str_SQL = str_SQL & " SELECT " & str_Fild
    str_SQL = str_SQL & " FROM [" & str_StRng & "] " '�Ǘ��\$A8:FZ8000
    str_SQL = str_SQL & " WHERE " & str_Key & " IS NOT NULL"
    Debug.Print str_SQL
    If str_WHERE <> "" Then
        str_SQL = str_SQL & str_WHERE
    End If
    Exl_Rs.Open str_SQL, Exl_Cn, adOpenKeyset, adLockReadOnly

End Function

Public Function Opn_AcRs(ByVal str_Tbl As String, str_Key As String, _
                                        Optional str_WHERE As String = "", _
                                        Optional str_Fild As String = "*", _
                                        Optional Flg As Long = 0)
'��Access���R�[�h�Z�b�g�I�[�v��
'(�����P:�e�[�u����,�����Q:Null���O�t�B�[���h���A�����R:�ǉ�������=�ȗ���""�A�����S:�t�B�[���h�w��=�ȗ���"*"�j
    Set Ac_Cn = New ADODB.Connection
    Set Ac_Rs = New ADODB.Recordset
    str_AcDBcn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                          Sheets("�f�B���N�g���ݒ�").Range("F8").Value & ";" 'AccessDB�p�X�̎擾�Ɛݒ�
'    On Error GoTo Era
    Ac_Cn.Open str_AcDBcn
    '��Access�f�[�^���oSQL
    str_SQL = ""
    str_SQL = str_SQL & " SELECT " & str_Fild
    str_SQL = str_SQL & " FROM " & str_Tbl
    If Flg = 0 Then
        str_SQL = str_SQL & " WHERE " & str_Key & " IS NOT NULL"
    End If
    If str_WHERE <> "" Then
        str_SQL = str_SQL & str_WHERE
    End If
    Debug.Print str_SQL
    Ac_Rs.Open str_SQL, Ac_Cn, adOpenForwardOnly, adLockPessimistic
    Exit Function
Era: '�G���[������*****************************************************
    If Err.Number = -2147467259 Then
        MsgBox "DB�t�@�C���֐ڑ��ł��܂���ł��� " & vbCrLf & _
         "�f�B���N�g���ݒ�Ńp�X���m�F�E�Đݒ肵�Ă�������" & vbCrLf & _
         "OK�������Ɛݒ�y�[�W�ֈړ����܂�", 16
         Call vis_SETDirectSt
         End
    Else
        MsgBox Err.Number & vbCrLf & _
         Err.Description, 16
         End
    End If

End Function

Public Function Dis_Exl_Rs()
'���Ǐo���R�[�h�Z�b�g�̃N���[�Y�Ɣj��
    On Error Resume Next
    If Exl_Rs Is Nothing Then
    Else
        Exl_Rs.Close
        Set Exl_Rs = Nothing
    End If
    If Exl_Cn Is Nothing Then
    Else
        Exl_Cn.Close
        Set Exl_Cn = Nothing
    End If
    
End Function

Public Function Dis_Ac_Rs()
'���Ǐo���R�[�h�Z�b�g�̃N���[�Y�Ɣj��
    On Error Resume Next
    If Ac_Rs Is Nothing Then
    Else
        Ac_Rs.Close
        Set Exl_Rs = Nothing
    End If
    If Ac_Cn Is Nothing Then
    Else
        Ac_Cn.Close
        Set Ac_Cn = Nothing
    End If
    
End Function
