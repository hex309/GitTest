Attribute VB_Name = "����_HTML�^�O"
Option Explicit


'-------------------------
'�Z���N�g�^�O
'-------------------------
Public Sub ieClickSelectBoxTagSelect(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)

    Call ieWaitCheck
    
    Dim �I�� As String

    �I�� = �I�v�V����1

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        
        Dim selectTAG As Object
        Dim optionList As Object
        Dim optionItem As Object

        For Each selectTAG In oIE.document.getElementsByTagName("select")

            If InStr(selectTAG.outerHTML, �N���b�N�Ώ�) = 0 Then GoTo continue2

            Set optionList = selectTAG.document.getElementsByName(�N���b�N�Ώ�)

            For Each optionItem In optionList.Item(0)
                
                If optionItem.innerText = �I�� Then

'                    selectTAG.selectedIndex = optionItem.Index
                    optionItem.Selected = True
                    selectTAG.onchange
                    Sleep 250
                    Exit Sub
                End If
            
            Next
            
continue2:
        Next
        
continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)

End Sub

'-------------------------
'�t�@�C���^���N���b�N
'-------------------------
Public Sub ieClickButtonTagInputTypeFile(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)
    
    Call ieWaitCheck
    
    Dim �������[�h1 As String, �������[�h2 As String
    Dim ���b�Z�[�W As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2
    ���b�Z�[�W = �e�L�X�g

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
         
            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) = 0 Then GoTo continue2
                
            If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
            If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2
            
            MessageBox 0, ���b�Z�[�W, "�m�F", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
            oTAG.Click
            Sleep 500
            Exit Sub

continue2:
        Next
        
continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�^�����������N�N���b�N
'-------------------------
Public Sub ieClickLinkTagAhrefTypeNone(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String)

    Call ieWaitCheck
    
    Dim �������[�h1 As String, �������[�h2 As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1

        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
        
            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) = 0 Then GoTo continue2
            
            If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
            If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2

            oTAG.Click
            Sleep 500
            Exit Sub

continue2:
        Next

continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�{�^���^���N���b�N
'-------------------------
Public Sub ieClickButtonTagInputTypeButton(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String)

    Call ieWaitCheck
    
    Dim �������[�h1 As String, �������[�h2 As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
         
            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) = 0 Then GoTo continue2
            
            If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
            If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2

            oTAG.Click
            Sleep 250
            Exit Sub

continue2:
        Next

continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�T�u�~�b�g�{�^���^���N���b�N
'-------------------------
Public Sub ieClickSubmitButtonTagInputTypeSubmit(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String)

    Call ieWaitCheck
    
    Dim �������[�h1 As String, �������[�h2 As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
         
            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) > 0 Then
                If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
                If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2
                
                oTAG.Click
                Sleep 250
                Exit Sub
            
            End If
continue2:
        Next
        
continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'���W�I�{�^���^���N���b�N
'-------------------------
Public Sub ieClickRadioButtonTagInputTypeRadio(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String)

    Call ieWaitCheck
    
    Dim �������[�h1 As String, �������[�h2 As String

    �������[�h1 = �I�v�V����1
    �������[�h2 = �I�v�V����2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
         
            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) > 0 Then
                If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
                If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2
                
                oTAG.Click
                Sleep 250
                Exit Sub
                
            End If
continue2:
        Next
                    
continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�`�F�b�N�{�b�N�X�^���N���b�N
'-------------------------
Public Sub ieClickCheckBoxTagInputTypeCheckBox(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String)

    Call ieWaitCheck
    
    Dim �^�U�l As Boolean
    Dim �C���f�b�N�X As Long

    �^�U�l = CBool(�I�v�V����1)
    �C���f�b�N�X = Val(�I�v�V����2)
    
    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin
            
        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1
            
        oIE.document.getElementsByName(�N���b�N�Ώ�)(�C���f�b�N�X).Checked = �^�U�l
        Sleep 250
        Exit Sub
        
continue1:
    Next

    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�e�L�X�g�{�b�N�X�^������
'-------------------------
Public Sub ieInTextBoxTagInputTypeText(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)
    
    Call ieWaitCheck
        
    Dim �C���f�b�N�X As Long
    
    �C���f�b�N�X = Val(�I�v�V����1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1

        oIE.document.getElementsByName(�N���b�N�Ώ�)(�C���f�b�N�X).Value = �e�L�X�g
        Sleep 250
        Exit Sub

continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'�p�X���[�h�^������
'-------------------------
Public Sub ieInPasswordBoxTagInputTypePassword(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)

    Call ieWaitCheck
    
    Dim �C���f�b�N�X As Long
    
    �C���f�b�N�X = Val(�I�v�V����1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1

        oIE.document.getElementsByName(�N���b�N�Ώ�)(�C���f�b�N�X).Value = �e�L�X�g
        Sleep 250
        Exit Sub
        
continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub

'-------------------------
'���o�F�B���^���̒l���o
'-------------------------
'Public Sub ieExValueTagInputTypeHidden(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)
'
'    Call ieWaitCheck
'
'    Dim �������[�h1 As String, �������[�h2 As String
'
'    �������[�h1 = �I�v�V����1
'    �������[�h2 = �I�v�V����2
'
'    Dim oIE As InternetExplorerMedium
'    Dim oShell As Object, oWin As Object
'    Dim bFLG As Boolean: bFLG = False
'
'    Set oShell = CreateObject("Shell.Application")
'
'    For Each oWin In oShell.Windows
'
'        If oWin.Name <> "Internet Explorer" Then GoTo continue1
'        Set oIE = oWin
'
'        If InStr(oIE.document.title, �E�B���h�E) = 0 Then GoTo continue1
'
'        Dim oTAG As Object
'        For Each oTAG In oIE.document.getElementsByTagName(�^�O)
'
'            If InStr(oTAG.outerHTML, �N���b�N�Ώ�) > 0 Then
'                If �������[�h1 <> "" And InStr(oTAG.outerHTML, �������[�h1) = 0 Then GoTo continue2
'                If �������[�h2 <> "" And InStr(oTAG.outerHTML, �������[�h2) = 0 Then GoTo continue2
'
'                Stop
'
'
'                Sleep 250
'                Exit Sub
'
'            End If
'continue2:
'        Next
'
'continue1:
'    Next
'
'    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
'
'
'End Sub

'-------------------------
'�e�L�X�g�G���A�^������
'-------------------------
Public Sub ieInTextTagTextAreaTypeText(ByVal �E�B���h�E As String, ByVal �A�N�V���� As String, ByVal �^�O As String, ByVal �N���b�N�Ώ� As String, ByVal �I�v�V����1 As String, ByVal �I�v�V����2 As String, ByVal �e�L�X�g As String)

    Call ieWaitCheck
        
    Dim �C���f�b�N�X As Long
    
    �C���f�b�N�X = Val(�I�v�V����1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, �E�B���h�E) = 0 Then GoTo continue1

        oIE.document.getElementsByName(�N���b�N�Ώ�)(�C���f�b�N�X).Value = �e�L�X�g
        Sleep 250
        Exit Sub

continue1:
    Next
    
    If bFLG = False Then Call IE�s���S����G���[(�E�B���h�E)
    
End Sub
