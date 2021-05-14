Attribute VB_Name = "xCtrl"
Option Explicit
Option Private Module

'============================================================
'　処理高速化制御限定モジュール(クラスの代わり)
'============================================================
'ピボットで重たいので全止めします

Private xlcSetting      As XlCalculation
Private blnCBSSetting   As Boolean

'処理落ち時始末
Sub 画面戻し()
    Call SubCtrlMovableCmd(ueppEnd)
End Sub

'前準備・後片付け（画面更新、マウスポインタ）
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

