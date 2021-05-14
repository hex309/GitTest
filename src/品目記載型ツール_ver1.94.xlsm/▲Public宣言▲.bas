Attribute VB_Name = "��Public�錾��"
Option Explicit

'----------------------------------
'�X���[�v�y�у��b�Z�[�W�{�b�N�X�iAPI�錾�j
'----------------------------------
#If VBA7 And Win64 Then
    Public Declare PtrSafe Sub Sleep Lib "KERNEL32" (ByVal ms As LongPtr)
    Public Declare PtrSafe Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As LongPtr, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#Else
    Public Declare Sub Sleep Lib "KERNEL32" (ByVal ms As Long)
    Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
#End If

Public Const MB_TOPMOST = 262144 ' &H40000
Public Const MB_OK = 0 ' &H0
Public Const MB_ICONINFOMATION = 64
Public Const MB_EXCLAMATION = 48


'----------------------------------
' �t�H�[����ʍőO�ʁiAPI�錾�j
'----------------------------------
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

#If VBA7 And Win64 Then
    Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

'----------------------------------------------------
'�p�u���b�N�N���X
'----------------------------------------------------
Public PubClsLucasAuth As LucasAuth:

'----------------------------------
'��~�t���O(StopButton�t�H�[��)
'----------------------------------
Public fStop As Boolean

Public Const IS_TEST As Boolean = False
'----------------------------------------------------
'�p�u���b�NIE�I�u�W�F�N�g
'----------------------------------------------------
Public oPubIE1  As InternetExplorerMedium

Public Pub�V�[�g�� As String
Public Pub����ԍ� As Long

'----------------------------------------------------
'�I�[�g�p�C���b�g�����ϐ��i���ϓo�^�ҁj
'----------------------------------------------------
Public Pub�I�[�g�p�C���b�g�ԍ� As String
Public Pub����V�[�g�� As String
Public Pub���σV�[�g�� As String

Public Pub���ϓo�^���� As String
Public Pub�c�Ǝ҃R�[�h As String
Public Pub��C�҃R�[�h As String

Public Pub�H��FROM As Variant
Public Pub�H��TO As Variant

Public Pub�X�܃R�[�h As String     '����_�H�������d�l���p
Public Pub�H�������d�l�� As String '����_�H�������d�l���p

Public Pub���ϑO����� As String '����_���ϓo�^�A�b�v���[�h�p

'----------------------------------------------------
'�V�[�g��
'----------------------------------------------------
Public Const WSNAME_TOOL As String = "�i�ڋL���^ (�c�[��)"
Public Const WSNAME_VAL_HINMOKU As String = "�ϐ��ꗗ_�i�ڊǗ�"
Public Const WSNAME_VAL_MODULE As String = "�ϐ��ꗗ_���W���[��"
Public Const WSNAME_LOG As String = "���s����"
Public Const WSNAME_LOGIN As String = "���O�C��"
Public Const WSNAME_CSVUP As String = "CSV�A�b�v���[�g"
Public Const WSNAME_WARIKOMI As String = "����T01"
Public Const WSNAME_CONFIG As String = "�ݒ�"
Public Const WSNAME_HINMOKU_MODULE As String = "�i�ڋL���^�i���W���[���j"
Public Const WSNAME_HINMOKU As String = "�i�ڊǗ��\"
Public Const WSNAME_CODE As String = "�R�[�h���X�g"
Public Const WSNAME_FORMAT As String = "���ϊm�F�t�H�[�}�b�g"

'----------------------------------------------------
'���̓t�@�C���@�}�X�^�[�t�@�C���֘A�i����\���ꗗ�t�@�C���j
'----------------------------------------------------
Public Pub�}�X�^�u�b�N�t���p�X As Variant 'C:\xxx\yyy\zzz.xlsx
Public Pub�}�X�^�u�b�N�p�X As String      'C:\xxx\yyy\
Public Pub�}�X�^�u�b�N�� As String        'zzz.xlsx

'----------------------------------------------------
'�o�̓t�@�C���@��Ǝ菇���t�@�C���֘A
'----------------------------------------------------
Public Pub�u�b�N�t���p�X As String 'C:\xxx\yyy\zzz.xlsx
'Public Pub�u�b�N�p�X As String      'C:\xxx\yyy\ 'Pub��ƃt�H���_�p�X�ϐ������邩�疢�g�p
Public Pub�u�b�N�� As String        'zzz.xlsx
Public Pub�g���q As String          'xlsx

'----------------------------------------------------
'�����t�@�C���@�}�X�^�[�t�@�C���̕����i�i����\���ꗗ�t�@�C���j
'----------------------------------------------------
'Public Pub�����u�b�N�t���p�X As Variant 'C:\xxx\yyy\zzz.xlsx
'Public Pub�����u�b�N�p�X As String      'C:\xxx\yyy\
'Public Pub�����u�b�N�� As String        'zzz.xlsx

'----------------------------------------------------
'���� �o�̓t�@�C�����ɓ����t�^�Ŏg�p
'----------------------------------------------------
Public YYYYMMDD_HHNNSS '�N����_�����b
Public YYYYMMDD        '�N����
Public HHNNSS          '�����b

'----------------------------------------------------
'�o�̓t�@�C���̍�Ɨp�t�H���_�p�X
'----------------------------------------------------
Public Pub��ƃt�H���_�p�X As String 'C:\xxx\yyy\

'----------------------------------------------------
'�V�K���R�[�h�t���O
'----------------------------------------------------
'Public Pub�V�K���_�� As Boolean
'Public Pub�V�K���i�� As Boolean

Public Pub�{���A�h���X As String
