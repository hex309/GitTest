VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} C01_UF01 
   Caption         =   "カレンダー"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "C01_UF01.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "C01_UF01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private PriColCommandButton As New Collection 'コマンドボタン

'「今月」ボタン
Private Sub CurrentMonthButton_Click()

    Dim MM As Long

    MM = Format(Now(), "MM")

    ComboBox1.ListIndex = 1
    ComboBox2.ListIndex = MM - 1

End Sub

'スピンボタン（ダウン）
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

    Call カレンダー生成

End Sub

'スピンボタン（アップ）
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

    Call カレンダー生成

End Sub

'フォームアクティブ
Sub UserForm_Activate()

    ComboBox1.SetFocus

End Sub

'フォーム初期化
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
        .AddItem YYYY - 1 & "年"
        .AddItem YYYY + 0 & "年"
        .AddItem YYYY + 1 & "年"
    End With

    ComboBox1.ListIndex = 1

    Me.ComboBox2.Clear

    With Me.ComboBox2
        .AddItem "1月"
        .AddItem "2月"
        .AddItem "3月"
        .AddItem "4月"
        .AddItem "5月"
        .AddItem "6月"
        .AddItem "7月"
        .AddItem "8月"
        .AddItem "9月"
        .AddItem "10月"
        .AddItem "11月"
        .AddItem "12月"
    End With

    ComboBox2.ListIndex = MM - 1

    Call カレンダー生成

    Application.EnableEvents = True

End Sub

'年
Sub ComboBox1_click()

    If Application.EnableEvents = False Then Exit Sub

    Call カレンダー生成

End Sub

'月
Sub ComboBox2_click()

    If Application.EnableEvents = False Then Exit Sub

    Call カレンダー生成

End Sub

'カレンダー
Sub カレンダー生成()

    Dim YYYY As Long
    Dim MM As Long

    YYYY = Replace(ComboBox1.Value, "年", "")
    MM = Replace(ComboBox2.Value, "月", "")

    Dim 月初 As Date
    Dim 月末 As Date
    Dim 処理日 As Date
    Dim 週目 As Long
    Dim 曜日 As Long

    月初 = DateSerial(YYYY, MM, 1)
    月末 = DateSerial(YYYY, MM + 1, 1) - 1

    曜日 = Weekday(月初)

    Dim i As Long
    For i = 1 To 42
        Controls("CommandButton" & i).TabStop = True
        Controls("CommandButton" & i).Locked = False
        Controls("CommandButton" & i).Caption = ""
    Next

    For i = 1 To Day(月末)
        処理日 = DateSerial(YYYY, MM, i)
        週目 = WorksheetFunction.WeekNum(処理日) - WorksheetFunction.WeekNum(月初) '1週目は0
        Controls("CommandButton" & i - 1 + 曜日).Caption = i
    Next

    For i = 1 To 42
        If Controls("CommandButton" & i).Caption = "" Then
            Controls("CommandButton" & i).TabStop = False
            Controls("CommandButton" & i).Locked = True
        End If
    Next

End Sub


