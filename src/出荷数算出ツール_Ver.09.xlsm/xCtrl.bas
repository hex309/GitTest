Attribute VB_Name = "xCtrl"
Option Explicit
Option Private Module

'============================================================
'�@����������������胂�W���[��(�N���X�̑���)
'============================================================
'�s�{�b�g�ŏd�����̂őS�~�߂��܂�

Private xlcSetting      As XlCalculation
Private blnCBSSetting   As Boolean

'�����������n��
Sub ��ʖ߂�()
    Call SubCtrlMovableCmd(ueppEnd)
End Sub

'�O�����E��Еt���i��ʍX�V�A�}�E�X�|�C���^�j
Public Sub SubCtrlMovableCmd(ByVal ueppStartEnd As uePrePost)

    If ueppStartEnd = ueppStart Then

        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
            .EnableEvents = False
            '            xlcSetting = .Calculation
            '            .Calculation = xlCalculationManual
            '            blnCBSSetting = .CalculateBeforeSave
            '            .CalculateBeforeSave = False
            .CutCopyMode = False
        End With

    ElseIf ueppStartEnd = ueppEnd Then

        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
            .EnableEvents = True
            '            .CalculateBeforeSave = blnCBSSetting
            '            .Calculation = xlcSetting
            .CutCopyMode = False
        End With

    End If

End Sub

