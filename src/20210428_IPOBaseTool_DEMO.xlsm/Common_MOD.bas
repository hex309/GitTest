Attribute VB_Name = "Common_MOD"
Option Explicit
'�����p���W���[��
'�����R�[�h�Z�b�g�擾�֘A�ϐ�
Public Const adOpenKeyset = 1, adLockReadOnly = 1
Public Exl_Cn, Ac_Cn As ADODB.Connection
Public Exl_Rs, Ac_Rs As ADODB.Recordset
Public Ac_Cmd As ADODB.Command
Public str_SQL, str_AcDBcn As String
'���u�b�N�N���[�Y�p�t���O
Public ClsFlg As Boolean

Public Sub Auto_open()
'���N��������
    ClsFlg = False
    Call St_AllUnvis
    ActiveSheet.Unprotect
    ActiveSheet.Range("C2").Select
    Call St_Lock
    Application.WindowState = xlMaximized   '�E�B���h�E�ő剻

End Sub
