Attribute VB_Name = "Get_SQL_MOD"
Option Explicit
'��SQL���ɕK�v�ȕ������擾���`���W���[��

Public Function Get_WHERE(ByVal str_Stn As String, str_FC As String, str_VC As String) As String
'��SQL�ǉ�WHERE��̍쐬
    '(����1:�擾�������V�[�g��,�����Q:�����t�B�[���h�̐擪�s�Z���A�h���X,����3:�����l�̐擪�s�Z���A�h���X)
    Dim i As Long
    Dim str_Ans As String
    Dim rs As ADODB.Recordset
    
    Application.ScreenUpdating = False
    Call Get_WHERELis(str_Stn, "T_WHEREList", str_FC, 1)
    Call Get_WHERELis(str_Stn, "T_WHEREList", str_VC, 2)
    Call Opn_ExlRs("T_WHEREList$A1:B200", "F1", , , 1)
    With Exl_Rs
        Do Until .EOF
            str_Ans = str_Ans & " AND " & !F1 & " Like '%" & !F2 & "%'"
            .MoveNext
        Loop
    End With
    Call Dis_Exl_Rs
    Get_WHERE = str_Ans
    
End Function

Public Function Get_SQLFelds(ByVal str_RStn As String) As String
'��SQL���t�B�[���h�w�蕔���擾
    '�t�B�[���h����J�����w�蕔���𐶐�
    Dim R_Ws As Worksheet
    Dim i, eCol, eRow As Long
    Dim str_Ans As String
    
    Set R_Ws = Sheets(str_RStn)
    With R_Ws
        str_Ans = ""
        eRow = .Cells(Rows.Count, 1).End(xlUp).Row
        For i = 1 To eRow
            str_Ans = str_Ans & .Cells(i, 1).Value & ","
        Next i
    End With
    str_Ans = Left(str_Ans, Len(str_Ans) - 1)
    Get_SQLFelds = str_Ans

End Function
