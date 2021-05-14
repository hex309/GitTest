VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} C01_UF01 
   Caption         =   "�J�����_�["
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "C01_UF01.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "C01_UF01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private PriColCommandButton As New Collection '�R�}���h�{�^��

'�u�����v�{�^��
Private Sub CurrentMonthButton_Click()

    Dim MM As Long

    MM = Format(Now(), "MM")

    ComboBox1.ListIndex = 1
    ComboBox2.ListIndex = MM - 1

End Sub

'�X�s���{�^���i�_�E���j
Private Sub SpinButton1_SpinDown()

    Dim yIndex As Long
    Dim mIndex As Long

    yIndex = ComboBox1.ListIndex
    mIndex = ComboBox2.ListIndex

    If mIndex > 0 Then
        ComboBox2.ListIndex = mIndex - 1
    Else
        If yIndex >= 1 Then
            ComboBox1.ListIndex = yIndex - 1
            ComboBox2.ListIndex = 11
        End If
    End If

    Call �J�����_�[����

End Sub

'�X�s���{�^���i�A�b�v�j
Private Sub SpinButton1_SpinUp()

    Dim yIndex As Long
    Dim mIndex As Long

    yIndex = ComboBox1.ListIndex
    mIndex = ComboBox2.ListIndex

    If mIndex < 11 Then
        ComboBox2.ListIndex = mIndex + 1
    Else
        If yIndex <= 1 Then
            ComboBox1.ListIndex = yIndex + 1
            ComboBox2.ListIndex = 0
        End If
    End If

    Call �J�����_�[����

End Sub

'�t�H�[���A�N�e�B�u
Sub UserForm_Activate()

    ComboBox1.SetFocus

End Sub

'�t�H�[��������
Sub Userform_initialize()

    Application.EnableEvents = False

    Dim myCtrl As Control
    Dim Cls As Object

    For Each myCtrl In Me.Controls
        If TypeName(myCtrl) = "CommandButton" Then
            Set Cls = New Class_C01_Commandbutton
            Cls.SetCtrl myCtrl
            PriColCommandButton.Add Cls
            Set Cls = Nothing
        End If
    Next

    Dim YYYY As Long
    Dim MM As Long

    YYYY = Format(Now(), "YYYY")
    MM = Format(Now(), "MM")

    Me.ComboBox1.Clear

    With Me.ComboBox1
        .AddItem YYYY - 1 & "�N"
        .AddItem YYYY + 0 & "�N"
        .AddItem YYYY + 1 & "�N"
    End With

    ComboBox1.ListIndex = 1

    Me.ComboBox2.Clear

    With Me.ComboBox2
        .AddItem "1��"
        .AddItem "2��"
        .AddItem "3��"
        .AddItem "4��"
        .AddItem "5��"
        .AddItem "6��"
        .AddItem "7��"
        .AddItem "8��"
        .AddItem "9��"
        .AddItem "10��"
        .AddItem "11��"
        .AddItem "12��"
    End With

    ComboBox2.ListIndex = MM - 1

    Call �J�����_�[����

    Application.EnableEvents = True

End Sub

'�N
Sub ComboBox1_click()

    If Application.EnableEvents = False Then Exit Sub

    Call �J�����_�[����

End Sub

'��
Sub ComboBox2_click()

    If Application.EnableEvents = False Then Exit Sub

    Call �J�����_�[����

End Sub

'�J�����_�[
Sub �J�����_�[����()

    Dim YYYY As Long
    Dim MM As Long

    YYYY = Replace(ComboBox1.Value, "�N", "")
    MM = Replace(ComboBox2.Value, "��", "")

    Dim ���� As Date
    Dim ���� As Date
    Dim ������ As Date
    Dim �T�� As Long
    Dim �j�� As Long

    ���� = DateSerial(YYYY, MM, 1)
    ���� = DateSerial(YYYY, MM + 1, 1) - 1

    �j�� = Weekday(����)

    Dim i As Long
    For i = 1 To 42
        Controls("CommandButton" & i).TabStop = True
        Controls("CommandButton" & i).Locked = False
        Controls("CommandButton" & i).Caption = ""
    Next

    For i = 1 To Day(����)
        ������ = DateSerial(YYYY, MM, i)
        �T�� = WorksheetFunction.WeekNum(������) - WorksheetFunction.WeekNum(����) '1�T�ڂ�0
        Controls("CommandButton" & i - 1 + �j��).Caption = i
    Next

    For i = 1 To 42
        If Controls("CommandButton" & i).Caption = "" Then
            Controls("CommandButton" & i).TabStop = False
            Controls("CommandButton" & i).Locked = True
        End If
    Next

End Sub


