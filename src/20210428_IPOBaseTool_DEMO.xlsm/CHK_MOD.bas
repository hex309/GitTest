Attribute VB_Name = "CHK_MOD"
Option Explicit
'���`�F�b�N�n���W���[��
Public Function CHK_Duplicate_ID(ByVal str_Tbln As String, _
                         str_Sval As String, str_Feldn As String) As Boolean
'���l�d���`�F�b�N
    Call Opn_AcRs(str_Tbln, str_Feldn, , str_Feldn)
    With Ac_Rs
        Do Until .EOF
            If str_Sval = Ac_Rs(str_Feldn).Value Then GoTo Skip
            .MoveNext
        Loop
    End With
    Call Dis_Ac_Rs
    CHK_Duplicate_ID = False
    Debug.Print CHK_Duplicate_ID
    Exit Function
Skip:
    CHK_Duplicate_ID = True
    Debug.Print CHK_Duplicate_ID
    Call Dis_Ac_Rs
    
End Function
Sub ttttteeetttst1()
    Debug.Print CHK_Duplicate_ID("T_KANRI", "XXX190201003", "T_1")
End Sub

Public Sub CHK_RegChange(Optional ByVal str_Stn As String = "�Ǘ��\�ҏW�o�^")
'���ύX�f�[�^�L���`�F�b�N�@�Ǘ��\�ҏW_�o�^�V�[�g��T_KANRI
    'DB��r�ŕύX���ꂽ���̂ɃV�[�g�J�����X�V�L��='�L' �Ǘ��\�㏑�X�V�����O�Ɏg�p
    Dim i, r, eRow As Long
    Dim str_ID, str_Fildn As String
    Dim AcVal, ExlVal As Variant
    Dim str_RngAd As String
    
    With Sheets(str_Stn)
        eRow = .Cells(Rows.Count, 4).End(xlUp).Row
        If eRow > 40 Then '�`�F�b�N�ʂ������ƕs����ɂȂ��(50���炢�����E�ۂ��j
            MsgBox "���R�[�h�����������܂�" & vbCrLf & "���R�[�h����30���ȓ��ɍi�����Ă�������", 16
            End
        End If
        For i = 10 To eRow
            str_ID = .Cells(i, 4).Value
            Call Opn_AcRs("T_KANRI", "T_1", " AND T_1='" & str_ID & "'")
            str_RngAd = Sheets(str_Stn).Range("B7").End(xlToRight).Address
            str_RngAd = Replace(str_RngAd, "7", "50")
            str_RngAd = Replace(str_RngAd, "$", "")
            Call Opn_ExlRs(str_Stn & "$D7:" & str_RngAd, "T_1", " AND T_1='" & str_ID & "'")
            For r = 3 To Exl_Rs.Fields.Count - 1
                str_Fildn = Exl_Rs.Fields(r).Name
                AcVal = IIf(IsNull(Ac_Rs(str_Fildn).Value), "", Ac_Rs(str_Fildn).Value)
                ExlVal = IIf(IsNull(Exl_Rs(str_Fildn).Value), "", Exl_Rs(str_Fildn).Value)
                If AcVal <> ExlVal Then
                    .Unprotect
                    .Cells(i, 2) = "�L"
                    GoTo Skip
                End If
            Next r
Skip:
            Call Dis_Exl_Rs
            Call Dis_Ac_Rs
        Next i
    End With
        
End Sub

Public Function CHK_WFildsNam(ByVal str_Stn As String, Coln As Long, _
                                                    sRow As Long, chk_Val As String) As Boolean
'���V�[�g�̒l�d���`�F�b�N
'(����1:�V�[�g��,����2:�J������No,����3:�擪�sNo,����4:�����l)
    Dim eRow, i As Long
    Dim Ws As Worksheet
    
    Set Ws = Sheets(str_Stn)
    With Ws
        eRow = .Cells(Rows.Count, Coln).End(xlUp).Row
        For i = sRow To eRow
            If .Cells(i, Coln) = chk_Val Then GoTo Skip
        Next i
    End With
    CHK_WFildsNam = False
    Exit Function
Skip:
    CHK_WFildsNam = True
 
End Function
