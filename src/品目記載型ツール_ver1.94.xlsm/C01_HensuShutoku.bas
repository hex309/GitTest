Attribute VB_Name = "C01_HensuShutoku"
Option Explicit
Option Private Module

'URL�Ȃǂ̒l�����[�N�V�[�g����擾
Private Sub �ϐ��擾�e�X�g()
    Debug.Print �ݒ�擾("���i�ڋL���^", "SSIS URL")
    Debug.Print �ݒ�擾("�����R�s�[�^", "SSIS URL")
End Sub
'URL�Ȃǂ̒l�����[�N�V�[�g����擾
Public Function �ݒ�擾(ByVal �Ώ� As String _
    , ByVal ���� As String) As Variant
    Dim sh�ݒ� As Worksheet
    Set sh�ݒ� = ThisWorkbook.Worksheets(WSNAME_CONFIG)
    
    Dim �Ώۗ� As Range
    With sh�ݒ�
        Set �Ώۗ� = .Rows(1).Find(�Ώ�)
    End With
    
    Dim oResult As Range
    With sh�ݒ�
        Set oResult = .Columns(�Ώۗ�.Column).Find(����)
    End With
    If oResult Is Nothing Then
        �ݒ�擾 = vbNullString
    Else
        �ݒ�擾 = oResult.Offset(0, 1).Value
    End If
End Function
