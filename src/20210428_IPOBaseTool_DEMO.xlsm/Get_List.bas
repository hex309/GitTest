Attribute VB_Name = "Get_List"
Option Explicit
'�����X�g�쐬�n���W���[��

Public Sub Get_KANRIColList()
'���Ǘ��\�f�[�^�t�H�[�}�b�g�@�t�B�[���h���̃��X�g�쐬
    '�J�����ݒ�Ǘ��\�J����ID���̓t�H�[���Ŏg�p
   Call Get_FildLis("�Ǘ��\�t�B�[���h�ݒ�$B5:GS6", "T_KANRIColList", "�Ǘ��\ID", 2)

End Sub

Public Sub Get_GAIBColList()
'���O���f�[�^�t�H�[�}�b�g�@�t�B�[���h���̃��X�g�쐬
    '�J�����ݒ�O���J����ID���̓t�H�[���Ŏg�p
        '255�ȏ��z��@Get_FildLis�Ŏ擾�ł���̂�MAX255�܂ł̈�
    Dim R_Ws As Worksheet
    Dim L_Ws As Worksheet
    Dim i, r As Long
    Dim CVal As String
    
    With ThisWorkbook
        Set R_Ws = .Sheets("T_GAIBCol")
        Set L_Ws = .Sheets("T_GAIBColList")
        L_Ws.Unprotect
        L_Ws.Range("B2:B500").ClearContents
        For i = 1 To 400
            CVal = R_Ws.Cells(1, i).Value
'            '�������s����ێ��������ꍇ�̓R�����g�A�E�g
'            If InStr(CVal, vbLf) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbLf, "")
'            ElseIf InStr(CVal, vbCr) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbCr, "")
'            ElseIf InStr(CVal, vbCrLf) > 0 Then
'                CVal = Replace(R_Ws.Cells(1, i).Value, vbCrLf, "")
'            End If
            L_Ws.Cells(i, 2).Value = CVal
        Next i
        .Save
    End With

End Sub

Public Function Get_FildLis(ByVal str_RStn As String, str_LStn As String, str_Nullkey As String, Colnam As Long)
'���t�B�[���h���X�g�̍쐬
    '�V�[�g�̃t�B�[���h�s�l���w��V�[�g��֏c�]�L
    '(����1:�擾�������V�[�g�͈�,����2:�]�L���������X�g�V�[�g��,����:3Null���O�t�B�[���h������4:�]�L��������ԍ�)
    Dim L_Ws As Worksheet
    Dim i As Long

    Application.ScreenUpdating = False
    Set L_Ws = Sheets(str_LStn)
    L_Ws.Unprotect
    L_Ws.Columns(Colnam).Clear
    Call Opn_ExlRs(str_RStn, str_Nullkey)
    With Exl_Rs
        For i = 0 To .Fields.Count - 1
            L_Ws.Cells(i + 1, Colnam).Value = Exl_Rs.Fields(i).Name
        Next i
    End With
    Call Dis_Exl_Rs
    
End Function

Public Function Get_WHERELis(ByVal str_RStn As String, str_LStn As String, str_FCAd As String, Colnam As Long)
'��WHERE�僊�X�g�̍쐬
    '�V�[�g�̃t�B�[���h�s�l���w��V�[�g��֏c�]�L
    '(����1:�擾�������V�[�g��,����2:�]�L���������X�g�V�[�g��,����:3�擪�t�B�[���h�Z���A�h���X,����4:�]�L��������ԍ�)
    Dim L_Ws As Worksheet
    Dim R_Ws As Worksheet
    Dim i, sRow, sCol As Long

    Application.ScreenUpdating = False
    Set R_Ws = Sheets(str_RStn)
    Set L_Ws = Sheets(str_LStn)
    With R_Ws
        sRow = .Range(str_FCAd).Row
        sCol = .Range(str_FCAd).Column
        L_Ws.Columns(Colnam).Clear
        For i = sCol To 200
            L_Ws.Cells(i - sCol + 1, Colnam).Value = .Cells(sRow, i).Value
        Next i
    End With

End Function
