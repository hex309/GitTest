Attribute VB_Name = "Exl_MOD"
Option Explicit
'��Excel��Excel�f�[�^����n���W���[��
Dim str_AcDBcn As String

Public Sub Imp_StD()
'���O���f�[�^�t�@�C��(�V�[�g)�I�[�v���˃f�[�^�擾�l�\��t��
    Dim Ws_UI As Worksheet
    Dim Ws_S As Worksheet
    Dim Ws_R As Worksheet
    Dim Ws_L As Worksheet
    Dim Ws_F As Worksheet
    Dim str_FPath As String
    Dim str_Stn As String
    Dim str_rng As String
    Dim str_FRng As String
    Dim eColn As Long
    
    With ThisWorkbook
        Set Ws_UI = .Sheets("�C���|�[�g")
        Set Ws_S = .Sheets("�O���f�[�^�V�[�g�͈͐ݒ�")
        Set Ws_L = .Sheets("�ꊇ�捞")
        Set Ws_F = .Sheets("T_GAIBCol")
    End With
    Application.ScreenUpdating = True
    With Ws_UI
        .Unprotect
        .Shapes("Fil_1").Visible = True
        .Shapes("Gr_1").Visible = True
    End With
    Application.Wait Now() + TimeValue("00:00:01")
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Call St_Lock
   
    str_FPath = Ws_UI.Range("C7").Value
    
    With Ws_S
        str_Stn = .Range("D8").Value
        str_rng = .Range("F8").Value & ":" & .Range("H8")
        str_FRng = .Range("J8").Value
        eColn = Mid(str_FRng, 2, 2)
        str_FRng = str_FRng & ":OZ" & eColn
    End With

    If str_FPath = "" Then
        MsgBox "�Ǎ��t�@�C���p�X��ݒ肵�Ă�������", 16
        End
    End If
    
    Ws_L.Range(str_rng).ClearContents
    On Error GoTo Era1
    Workbooks.Open str_FPath
    ActiveWorkbook.Sheets(str_Stn).Range(str_rng).Copy
    Ws_L.Range("A2").PasteSpecial Paste:=xlValues
    '�O���t�B�[���h���X�g���擾
    Ws_F.Unprotect
    Ws_F.Cells.ClearContents
    ActiveWorkbook.Sheets(str_Stn).Range(str_FRng).Copy
    Ws_F.Range("A1").PasteSpecial Paste:=xlValues
    
    ActiveWorkbook.Close
    ThisWorkbook.Save

    Exit Sub
Era1:
Debug.Print Err.Number
    If Err.Number = 1004 Then
        MsgBox "�y�V�t�@�C���֐ڑ��ł��܂���ł��� " & vbCrLf & _
         "�p�X���m�F�E�Đݒ肵�Ă�������", 16
         Call Unvis_Imp
         End
    End If

    MsgBox "�Ǎ��V�[�g�͈͂��������Ȃ��悤�ł�" & vbCrLf & _
                "�͈͂̐ݒ���m�F�E�Đݒ肵�Ă�������", 16
        Call Unvis_Imp
    End

End Sub

Public Function TMP_StD(ByVal str_RStn As String, str_LStn As String, _
                                        str_RRng As String, str_LRng As String)
'���y�V�f�[�^�V�[�g�ˈꎞ�e�[�u���V�[�g�֒l�\��t��
    '�@�@�����P�@�@�@�@�@�@�����Q�@�@�@�@�@�����R�@�@�@�@�����S
    '("�Ǐo�V�[�g��","�����V�[�g��","�Ǐo�V�[�g�͈�","�����V�[�g�Z��")
    '��("TMP1","TMP2","A1:Z:1000","A1")

    Dim Ws_R As Worksheet
    Dim Ws_L As Worksheet
    Dim Ws_LC As Worksheet
    Dim str_LCStn As String
    
    str_LCStn = Replace(str_LStn, "T_", "CHK_")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    With ThisWorkbook
        Set Ws_R = .Sheets(str_RStn)
        Set Ws_L = .Sheets(str_LStn)
        Set Ws_LC = .Sheets(str_LCStn)
    End With
    If str_LRng <> "C3" Then
        Ws_L.Range("A3:GR9000").ClearContents
        Ws_LC.Range("A3:GR9000").ClearContents
    Else
        Ws_L.Range("C3:GT9000").ClearContents
        Ws_LC.Range("C3:GT9000").ClearContents
    End If
    Ws_R.Range(str_RRng).Copy
    Ws_L.Range(str_LRng).PasteSpecial Paste:=xlValues
    Ws_LC.Range(str_LRng).PasteSpecial Paste:=xlValues
    ThisWorkbook.Save

End Function

Public Sub Imp_CSV_Exl()
'��CSV����Excel�ɂƂ肱��
    Dim Fn As Variant
    
    Application.ScreenUpdating = False
    Fn = Sheets("CSV�C���|�[�g").Range("C7").Value
    If Fn = "" Then
        Exit Sub
    End If
    Sheets("TMP_CSV").Unprotect
    Sheets("TMP_CSV").Cells.ClearContents
    Workbooks.Open FileName:=Fn
    ActiveSheet.Cells.Copy ThisWorkbook.Sheets("TMP_CSV").Cells
    ActiveWorkbook.Close SaveChanges:=False
    
End Sub
