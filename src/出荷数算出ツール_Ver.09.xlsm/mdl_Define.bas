Attribute VB_Name = "mdl_Define"
Option Explicit
Option Private Module

'============================================================
'   �O���[�o���萔�E�ėp�񋓌^
'============================================================


'----------------
'�@�萔
'----------------
'�t�@�C���ݒ�i���V�[�g���j
Public Const G_FILE_TARGET_OA   As String = "�����E���׎��шꗗ"
Public Const G_FILE_TARGET_SS   As String = "���i�䒠�i�����j"
Public Const G_FILE_TARGET_DD   As String = "���i�䒠�i����j"
Public Const G_FILE_TARGET_YA   As String = "�o�׎���"
Public Const G_FILE_TARGET_SU   As String = "�d����ꗗ"


'���ڃ��X�g���C���i�e�t�@�C�������ږ��j���\��t����̖��̂ɍ��킹�Ȃ��ŉ�����
Public Const G_ITEMLINE_OA      As String = "�o�ח\��,���i,����,�敪" 'X_�󒍈��������ږ�
Public Const G_ITEMLINE_SS      As String = "�s���x��,�ݒ萔" '���S�݌ɐ������ږ�


'�c�[���N���V�[�g
Public Const G_SHEETNAME_TOOL       As String = "���f�[�^�Ǎ�"

'�v�Z�p�V�[�g
Public Const G_SHEETNAME_CALCBASE   As String = "�v�Z�p"

'���C���V�[�g
Public Const G_SHEETNAME_MAIN       As String = "�Z�o�c�[��"


'�Z�o�p�s�{�b�g��
Public Const G_PIVOTNAME_MAIN       As String = "pivot�݌�"


'���C���V�[�g���W�J�^�[�Q�b�g
Public Const G_PAST_TARGET_PARTNUMBER   As String = "���i�ԍ�"
Public Const G_PAST_TARGET_CALCRESULT   As String = "���ݍ݌ɐ�"


'���S�݌ɐ��t�@�C�����̈��S�݌ɂ̗�ʒu
Public Const C_���S�݌ɐ���         As Long = 3



'�g���q
'�捞�p�b�r�u�t�@�C���̊g���q
Public Const G_EXT_CSV          As String = ".csv"

'�e�L�X�g�t�@�C���̊g���q
Public Const G_EXT_TXT          As String = ".txt"

'Excel�t�@�C���̊g���q
Public Const G_EXT_XLSX         As String = ".xlsx"
Public Const G_EXT_XLS          As String = ".xls*"

'�t�@�C���t�B���^�[�p������
Public Const G_FILTERNAME_ALL   As String = "���ׂẴt�@�C��(*.*)"
Public Const G_FILTERNAME_XLSX  As String = "Excel�u�b�N(*.xlsx)"
Public Const G_FILTERNAME_EXCEL As String = "Excel�t�@�C��(*.xls*)"
Public Const G_FILTERNAME_CSV   As String = "CSV�`��(*.csv)"

'������t�H�[�}�b�g�`��
Public Const G_FORMAT_DATE_YYYYMMDD         As String = "yyyymmdd"


'�G���[�n���h�����O�p�ݒ�
Public Const G_CTRL_ERROR_NUMBER_USER_NOTICE    As Long = vbObjectError + 1
Public Const G_CTRL_ERROR_NUMBER_USER_CAUTION   As Long = vbObjectError + 9
Public Const G_CTRL_ERROR_NUMBER_DEVELOPER      As Long = vbObjectError + 100



'----------------
'�@�񋓌^
'----------------
'�t�@�C���̎��
Public Enum ueXLFileType
    [_MIN] = 0
    uexlOrderArrival = 1                         '����_���׎���
    uexlPreProductBook = 2                       '���i�䒠�i�O���j
    uexlProductBook = 3                          '���i�䒠�i�����j
    uexlSupplierBook = 4                         '�d����
    [_MAX]
End Enum

Public Enum ueRC
    [_MIN] = 0
    uercRow = 1
    uercCol = 2
    [_MAX]
End Enum

'��ԍ��̐��l��
Public Enum ueColumnNum
    ueColA = 1
    ueColB = 2
    ueColC = 3
    ueColD = 4
    ueColE = 5
    ueColF = 6
    ueColG = 7
    ueColH = 8
    ueColI = 9
    ueColJ
    ueColK
    ueColL
    ueColM
    ueColN
End Enum

'��L��萔��
Public Const COL_A     As Long = ueColA
Public Const COL_B     As Long = ueColB
Public Const COL_C     As Long = ueColC
Public Const COL_D     As Long = ueColD
Public Const COL_E     As Long = ueColE
Public Const COL_F     As Long = ueColF
Public Const COL_G     As Long = ueColG
Public Const COL_H     As Long = ueColH
Public Const COL_I     As Long = ueColI
Public Const COL_J     As Long = ueColJ
Public Const COL_K     As Long = ueColK
Public Const COL_L     As Long = ueColL
Public Const COL_M     As Long = ueColM
Public Const COL_N     As Long = ueColN
