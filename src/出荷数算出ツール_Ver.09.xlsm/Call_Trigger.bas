Attribute VB_Name = "Call_Trigger"
Option Explicit

'============================================================
'   ���[�U���Ăׂ�}�N��
'============================================================

Public Sub ����_���׎��шꗗ�t�@�C���I��()
    Call SubSetCSVFilePathByUserChoice(uexlOrderArrival)
End Sub

Public Sub ���i�䒠_�O��_�t�@�C���I��()
    Call SubSetCSVFilePathByUserChoice(uexlPreProductBook)
End Sub

Public Sub ���i�䒠_����_�t�@�C���I��()
    Call SubSetCSVFilePathByUserChoice(uexlProductBook)
End Sub

Public Sub �ۑ���t�H���_�I��()
    Call GetFoldePath
End Sub

Public Sub �o�א��Z�o()
    Call GetCurrentData
End Sub

Public Sub �d����_�t�@�C���I��()
    Call SubSetCSVFilePathByUserChoice(uexlSupplierBook)
End Sub

Public Sub �d����ꗗ�擾()
    Call GetSupplier
End Sub

