Attribute VB_Name = "共通_HTMLタグ"
Option Explicit


'-------------------------
'セレクトタグ
'-------------------------
Public Sub ieClickSelectBoxTagSelect(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)

    Call ieWaitCheck
    
    Dim 選択名 As String

    選択名 = オプション1

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        
        Dim selectTAG As Object
        Dim optionList As Object
        Dim optionItem As Object

        For Each selectTAG In oIE.document.getElementsByTagName("select")

            If InStr(selectTAG.outerHTML, クリック対象) = 0 Then GoTo continue2

            Set optionList = selectTAG.document.getElementsByName(クリック対象)

            For Each optionItem In optionList.Item(0)
                
                If optionItem.innerText = 選択名 Then

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
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)

End Sub

'-------------------------
'ファイル型式クリック
'-------------------------
Public Sub ieClickButtonTagInputTypeFile(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)
    
    Call ieWaitCheck
    
    Dim 検索ワード1 As String, 検索ワード2 As String
    Dim メッセージ As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2
    メッセージ = テキスト

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(タグ)
         
            If InStr(oTAG.outerHTML, クリック対象) = 0 Then GoTo continue2
                
            If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
            If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2
            
            MessageBox 0, メッセージ, "確認", MB_OK Or MB_TOPMOST Or MB_ICONINFOMATION
            
            oTAG.Click
            Sleep 500
            Exit Sub

continue2:
        Next
        
continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'型式無しリンククリック
'-------------------------
Public Sub ieClickLinkTagAhrefTypeNone(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String)

    Call ieWaitCheck
    
    Dim 検索ワード1 As String, 検索ワード2 As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1

        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(タグ)
        
            If InStr(oTAG.outerHTML, クリック対象) = 0 Then GoTo continue2
            
            If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
            If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2

            oTAG.Click
            Sleep 500
            Exit Sub

continue2:
        Next

continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'ボタン型式クリック
'-------------------------
Public Sub ieClickButtonTagInputTypeButton(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String)

    Call ieWaitCheck
    
    Dim 検索ワード1 As String, 検索ワード2 As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(タグ)
         
            If InStr(oTAG.outerHTML, クリック対象) = 0 Then GoTo continue2
            
            If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
            If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2

            oTAG.Click
            Sleep 250
            Exit Sub

continue2:
        Next

continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'サブミットボタン型式クリック
'-------------------------
Public Sub ieClickSubmitButtonTagInputTypeSubmit(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String)

    Call ieWaitCheck
    
    Dim 検索ワード1 As String, 検索ワード2 As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(タグ)
         
            If InStr(oTAG.outerHTML, クリック対象) > 0 Then
                If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
                If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2
                
                oTAG.Click
                Sleep 250
                Exit Sub
            
            End If
continue2:
        Next
        
continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'ラジオボタン型式クリック
'-------------------------
Public Sub ieClickRadioButtonTagInputTypeRadio(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String)

    Call ieWaitCheck
    
    Dim 検索ワード1 As String, 検索ワード2 As String

    検索ワード1 = オプション1
    検索ワード2 = オプション2

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")
    
    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
        
        Dim oTAG As Object
        For Each oTAG In oIE.document.getElementsByTagName(タグ)
         
            If InStr(oTAG.outerHTML, クリック対象) > 0 Then
                If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
                If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2
                
                oTAG.Click
                Sleep 250
                Exit Sub
                
            End If
continue2:
        Next
                    
continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'チェックボックス型式クリック
'-------------------------
Public Sub ieClickCheckBoxTagInputTypeCheckBox(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String)

    Call ieWaitCheck
    
    Dim 真偽値 As Boolean
    Dim インデックス As Long

    真偽値 = CBool(オプション1)
    インデックス = Val(オプション2)
    
    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin
            
        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1
            
        oIE.document.getElementsByName(クリック対象)(インデックス).Checked = 真偽値
        Sleep 250
        Exit Sub
        
continue1:
    Next

    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'テキストボックス型式入力
'-------------------------
Public Sub ieInTextBoxTagInputTypeText(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)
    
    Call ieWaitCheck
        
    Dim インデックス As Long
    
    インデックス = Val(オプション1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1

        oIE.document.getElementsByName(クリック対象)(インデックス).Value = テキスト
        Sleep 250
        Exit Sub

continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'パスワード型式入力
'-------------------------
Public Sub ieInPasswordBoxTagInputTypePassword(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)

    Call ieWaitCheck
    
    Dim インデックス As Long
    
    インデックス = Val(オプション1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1

        oIE.document.getElementsByName(クリック対象)(インデックス).Value = テキスト
        Sleep 250
        Exit Sub
        
continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub

'-------------------------
'抽出：隠し型式の値抽出
'-------------------------
'Public Sub ieExValueTagInputTypeHidden(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)
'
'    Call ieWaitCheck
'
'    Dim 検索ワード1 As String, 検索ワード2 As String
'
'    検索ワード1 = オプション1
'    検索ワード2 = オプション2
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
'        If InStr(oIE.document.title, ウィンドウ) = 0 Then GoTo continue1
'
'        Dim oTAG As Object
'        For Each oTAG In oIE.document.getElementsByTagName(タグ)
'
'            If InStr(oTAG.outerHTML, クリック対象) > 0 Then
'                If 検索ワード1 <> "" And InStr(oTAG.outerHTML, 検索ワード1) = 0 Then GoTo continue2
'                If 検索ワード2 <> "" And InStr(oTAG.outerHTML, 検索ワード2) = 0 Then GoTo continue2
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
'    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
'
'
'End Sub

'-------------------------
'テキストエリア型式入力
'-------------------------
Public Sub ieInTextTagTextAreaTypeText(ByVal ウィンドウ As String, ByVal アクション As String, ByVal タグ As String, ByVal クリック対象 As String, ByVal オプション1 As String, ByVal オプション2 As String, ByVal テキスト As String)

    Call ieWaitCheck
        
    Dim インデックス As Long
    
    インデックス = Val(オプション1)

    Dim oIE As InternetExplorerMedium
    Dim oShell As Object, oWin As Object
    Dim bFLG As Boolean: bFLG = False

    Set oShell = CreateObject("Shell.Application")

    For Each oWin In oShell.Windows
    
        If oWin.Name <> "Internet Explorer" Then GoTo continue1
        Set oIE = oWin

        If InStr(oIE.document.Title, ウィンドウ) = 0 Then GoTo continue1

        oIE.document.getElementsByName(クリック対象)(インデックス).Value = テキスト
        Sleep 250
        Exit Sub

continue1:
    Next
    
    If bFLG = False Then Call IE不完全操作エラー(ウィンドウ)
    
End Sub
