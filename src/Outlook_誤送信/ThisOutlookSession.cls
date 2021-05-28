VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'// ---------------------------------------------------------------------------------------------------------------------
'//  Rev    ：1.6β1
'//  更新日 ：2019/10/2 64bit対応テスト用
'//         ：2019/10/25 添付ファイルがzipだった場合の対応用
'//         ：           添付ファイルの表示名が異なる場合のバグ対策版
'//         ：2020/10/09 同一パスワードが不定期ではあるが、生成されてしまうため
'//         ：           乱数の発生タイミングを変更
'//         ：           テストでは10,000件のパスワードを連続で生成しても重複はなし
'-------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Const MSG_VER As String = "メール誤送信対策Tool Ver1.5β1"

Private WithEvents olExp As Outlook.Explorer
Attribute olExp.VB_VarHelpID = -1

#If Win64 Then
  ' 64Bit 版
    Private Declare PtrSafe Function SevenZip Lib "7-zip64.DLL" ( _
                                  ByVal hWnd As Long, _
                                  ByVal szCmdLine As String, _
                                  ByVal szOutput As String, _
                                  ByVal dwSize As Long) As Long
#Else
  ' 32Bit 版
    Private Declare Function SevenZip Lib "7-zip32.DLL" ( _
                                  ByVal hWnd As Long, _
                                  ByVal szCmdLine As String, _
                                  ByVal szOutput As String, _
                                  ByVal dwSize As Long) As Long
#End If
                                  
Private Enum ZipState
    eTrue = 1
    eFalse = 2
    eRetry = 3
End Enum

#Const IsTest = False

Private Sub Application_Startup()
    Set olExp = Application.ActiveExplorer
    '#2020/10/09 起動時に初期化するように修正
    'パスワード生成時に行っていたが同一パスワードが作られてしまう
    'そのため生成時の初期化はやめ、起動時のみとする
    Randomize
End Sub

Private Sub olExp_InlineResponse(ByVal Item As Object)
    '返信用メールを表示（ポップアウト）」
    'この処理を入れないと、パスワード用メールを作る処理で、Copyメソッドが失敗する
    '（Copyメソッドは、インライン応答メールには対応していない）
    Item.Display
End Sub

'------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：Application_ItemSendイベント
'//  機能　　　　　　：メール送信時のイベント
'//  引数　　　　　　：
'//  戻り値　　　　　：なし
'//  作成日　　　　　：2015/09/10
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
    If CheckRecipients(Item) Then
        Cancel = True
        Exit Sub
    End If
    
    '添付ファイルの圧縮確認はマスト
'    Cancel = AttachCompressedZip(Item)
    Dim vResult As ZipState
    vResult = AttachCompressedZip(Item)
    If vResult = ZipState.eFalse Then
        Cancel = False
    ElseIf vResult = ZipState.eRetry Then
        Cancel = True
    ElseIf vResult = ZipState.eTrue Then
        Cancel = True
    Else
        'Do Nothing
    End If
    
    #If IsTest Then
        Cancel = True
    #End If
End Sub


'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：CheckRecipients
'//  機能　　　　　　：指定したドメインが宛先に含まれている場合、送信をキャンセルする
'//  引数　　　　　　：
'//  戻り値　　　　　：なし
'//  作成日　　　　　：2015/09/10
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Function CheckRecipients(ByVal Item As Object) As Boolean
    Dim CheckAddress As Variant
    Dim Recipient As Object
    Dim RecipientAddress As String

    Dim ErrDomain As String
    Dim ErrMsg As String
    Dim HasErr As Boolean
    Dim HasChecked As Boolean

    On Error GoTo ErrHdl

    'チェック対象ドメインの取得
    CheckAddress = GetCheckDomain

    '宛先アドレス（To、CC、BCC）の取得
    For Each Recipient In Item.Recipients
        RecipientAddress = ";" & GetNameAddrEx(Recipient) & RecipientAddress
    Next
    If Left$(RecipientAddress, 1) = ";" Then
        RecipientAddress = Mid$(RecipientAddress, 2)
    End If

    'チェック
    ErrDomain = HasErrDomain(RecipientAddress, CheckAddress)
    If Len(ErrDomain) = 0 Then
        'Do Nothing
    Else
        HasChecked = True
        ErrMsg = "このメッセージには以下のドメインを持つアドレスが含まれています" & vbCrLf _
               & "送信する場合は [OK] をクリックしてください" & vbCrLf _
               & ErrDomain
    End If

ExitHdl:
    If HasErr Then
        CheckRecipients = True
        Exit Function
    End If
    If HasChecked Then
        If MsgBox(ErrMsg, vbOKCancel + vbDefaultButton2 + vbCritical, MSG_VER) = vbOK Then
            If MsgBox("本当に送信しますか？", vbYesNo + vbDefaultButton2 + vbCritical, MSG_VER) = vbYes Then
                CheckRecipients = False
            Else
                CheckRecipients = True
            End If
        Else
            CheckRecipients = True
        End If
    End If
    Exit Function
ErrHdl:
    HasErr = True
    Resume ExitHdl
End Function


' ZIP 圧縮を行いメールを作成する VBA マクロ
Private Function AttachCompressedZip(ByVal objItem As Object) As ZipState
    Dim oFSO As Object    ' As Scripting.FileSystemObject
    Dim tempFolderPath As String
    Dim oFolder As Object    ' As Scripting.Folder
    Dim oNewMailItem As MailItem
    Dim tempItem As Object
    Dim AttachmentsCount As Long
    Dim i As Long
    Dim pos As Long
    Dim oAttach As Attachment
    Dim AttachFileName As String
    Dim TargetFileName As String
    Dim vExtention As String
    Dim ZipFileName As String
    Dim vPassword As String
    Dim RecipientAddress As String
    Dim msg As String
    Dim oRecipient As Object
    Dim tempAddress As String
    Dim recp As Recipient
    
    On Error GoTo ErrHdl
    ' 添付なければ終了
    For Each tempItem In objItem.Attachments
        '埋め込み画像かのチェック
        If IsAttachEmbedded(tempItem) = False Then
            AttachmentsCount = AttachmentsCount + 1
        End If
    Next
    If AttachmentsCount < 1 Then
        AttachCompressedZip = ZipState.eFalse
        Exit Function
    End If

    msg = "添付ファイルの暗号化が必要です" & vbCrLf _
        & "パスワード付きの圧縮をしますか？"
    If MsgBox(msg, vbYesNo, MSG_VER) = vbNo Then
        AttachCompressedZip = ZipState.eFalse
        Exit Function
    End If
    
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    ' 作業フォルダの作成 (%TEMP% で指定される一時フォルダの下にランダムな名前のフォルダを作成)
    tempFolderPath = oFSO.GetSpecialFolder(2) & "\" & oFSO.GetTempName()
    Set oFolder = oFSO.CreateFolder(tempFolderPath)

    ' パスワードを生成
    vPassword = MakePassword

    ' 作業フォルダに作成するファイルの名前を取得
    For Each tempItem In objItem.Attachments
        '埋め込み画像かのチェック
        If IsAttachEmbedded(tempItem) = False Then
            AttachFileName = tempFolderPath & "\" & tempItem.FileName
            pos = InStrRev(tempItem.FileName, ".")
            TargetFileName = Left(tempItem.FileName, pos - 1)
            vExtention = Right(tempItem.FileName, pos)
        End If
    Next

    '機種依存文字のチェック
    For Each tempItem In objItem.Attachments
        If InStr(tempItem.FileName, ChrW(8215)) > 0 Then
            MsgBox "ファイル名に機種依存文字が使用されています" & vbNewLine _
                & "ファイル名を確認してください", vbInformation
            AttachCompressedZip = ZipState.eRetry
            Exit Function
        End If
    Next
    
    '複数ファイルをまとめてZipにするため、ファイル名の指定を行う
    '2019/10/25 中林追加：添付ファイルが1つで、zipファイルの場合、ファイル名の指定を行う
    If AttachmentsCount > 1 _
        Or (AttachmentsCount = 1 And UCase(SplitPath(objItem.Attachments.Item(1).FileName)) = "ZIP") Then
        'ただし、デフォルトは最初に見つけたファイル名
        Do
            TargetFileName = InputBox("以下に zip ファイル名を入力してください", "ファイル名確認", TargetFileName)
            Select Case True
                Case StrPtr(TargetFileName) = 0
                    MsgBox "キャンセルされました", vbInformation, MSG_VER
                    AttachCompressedZip = ZipState.eTrue
                    Exit Function
                Case Len(Trim(TargetFileName)) = 0
                    MsgBox "エラー：ファイル名がありません" & vbCrLf _
                        & "指定してください", vbExclamation, MSG_VER
                        'Exit Function
                'ファイル名に使えない文字を指定したときの処理
                Case Not CheckFileName(Trim(TargetFileName))
                    MsgBox "ファイル名に指定できない文字があります" & vbCrLf _
                         & "再度指定して下さい", vbInformation, MSG_VER
                Case Else
                    Exit Do
            End Select
        Loop
    End If

    ZipFileName = oFSO.GetSpecialFolder(2) & "\" & TargetFileName & ".zip"

    ' 添付ファイルの一つ一つについてチェック
    'Countプロパティの値が正しくない場合があるので、エラー処理
    On Error Resume Next
    For i = objItem.Attachments.Count To 1 Step -1
        Set oAttach = objItem.Attachments.Item(i)
        If IsAttachEmbedded(oAttach) = False Then
            AttachFileName = tempFolderPath & "\" & oAttach.FileName
            oAttach.SaveAsFile AttachFileName
        End If
    Next
    On Error GoTo 0

    On Error GoTo ErrHdl
    ' 暗号化 zip ファイルを作成
    If Not SendToZIP(tempFolderPath & "\", ZipFileName, vPassword) Then
        MsgBox "添付ファイルの圧縮に失敗しました。ファイルを確認してください", , MSG_VER
        AttachCompressedZip = ZipState.eTrue
        GoTo ExitHdl
    End If

    '元の貼付ファイルを削除
    On Error Resume Next
    For i = objItem.Attachments.Count To 1 Step -1
        Set oAttach = objItem.Attachments.Item(i)
        If IsAttachEmbedded(oAttach) = False Then
            oAttach.Delete
        End If
    Next
    On Error GoTo 0
    
    On Error GoTo ErrHdl
    
    '2017/05/24 橋本 編集　'objItemがメール以外のケースを追加
    
    If objItem.Class = olMail Then
        '宛先等を引き継げるようにメールをコピーする
        Set oNewMailItem = objItem.Copy
    Else
        'objItemがメール以外だった場合、あて先を抜き出してメールを新規作成
        Set oNewMailItem = Application.CreateItem(olMailItem)
        
        For Each recp In objItem.Recipients
            oNewMailItem.Recipients.Add (GetNameAddrEx(recp))
        Next
        
    End If
    
    'アドレス帳と照合
    oNewMailItem.Recipients.ResolveAll
    
    ' 新規メール（アイテム）を作成
    oNewMailItem.Display
    
    ' zip 後のファイルを添付
    objItem.Attachments.Add ZipFileName, olByValue, , TargetFileName ' oAttach.DisplayName

    ' 先頭にパスワードに関する注意文を挿入する
    AddTextInBody objItem
    
    
    ' 新規メールの作成
    oNewMailItem.Subject = "[パスワード]" & objItem.Subject
    oNewMailItem.BodyFormat = olFormatPlain   '1
    oNewMailItem.Body = "**************************************************************" & vbCr & _
                        "　　添付 zip ファイルの復号パスワードをお知らせ致します。" & vbCr & _
                        "**************************************************************" & vbCr & vbCr & _
                        "先にお送りしたメールの添付ファイルのパスワードは以下のとおりです。" & vbCr & vbCr & _
                        "[メッセージ]" & vbCr & _
                        "題名:" & objItem.Subject & vbCr & vbCr & _
                        "[パスワード]" & vbCr & _
                        vPassword & vbCr & vbCr & _
                        "お手数をおかけしますが、ご確認のほどよろしくお願い致します。"
    On Error GoTo 0
ExitHdl:
    On Error Resume Next
    oFSO.DeleteFile ZipFileName
    ' 作業フォルダを削除
    oFolder.Delete
    On Error GoTo 0
    Exit Function
ErrHdl:
    MsgBox Err.Description, vbExclamation, MSG_VER
    AttachCompressedZip = ZipState.eTrue
    Resume ExitHdl
End Function

'パスワード送信メールの本文追加
Private Sub AddTextInBody(ByVal objMail As Object)
    Dim objDoc As Object
    Dim MailBody As String

    Set objDoc = objMail.GetInspector.WordEditor
    
    MailBody = "**************************************************************" & vbCr & _
               "添付 zip ファイルの復号パスワードは別メールにてご連絡致します。" & vbCr & _
               "**************************************************************" & vbCr & vbCr

    '先頭に文言を挿入（署名の手前にしたかったけど、署名の位置を正しく取得できないため）
    '通常であれば行番号を取得すればよいが、本文が長いと正しく取得できないことがある
    objDoc.Range(0, 0).Text = MailBody
    
    'ObjMailがMailItem以外（MeetingItem）の場合、不要なInspectorが上がるのでMailItemのときだけDisplay。2017/05/24 橋本
    'MailItemはDisplayしないとこの後の処理に問題が出るケースがある。
    
    If objMail.Class = olMail Then
        objMail.Display
    End If
End Sub

Private Sub MakePasswordTest()
    Dim xApp As Object
    Set xApp = CreateObject("Excel.Application")
    xApp.Visible = True
    Dim wb As Object
    Set wb = xApp.WorkBooks.Add
    
    Dim i As Long
    For i = 1 To 10000
        wb.Worksheets(1).Cells(i, 1).Value = "'" & MakePassword
    Next
End Sub
'パスワードの作成
Private Function MakePassword(Optional MaxLength As Long = 8)
    Dim tempPassword As String
    Dim LowerList As String
    Dim UpperList As String
    Dim NumberList As String
    Dim SignList As String
    Dim TempChars As String
    Dim MaxNum As Integer
    Dim i As Integer

    LowerList = "abcdefghijklmnopqrstuvwxyz"
    UpperList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    NumberList = "123456789"
    '    SignList = "-^[];:,./!@#$%^&*()_-+=""'~|`{}<>?"
    '記号は以下に限定 圧縮時のコマンドでエラーになる可能性あり
    SignList = "[];:!#$_+=|{}<>"

    TempChars = LowerList & NumberList & UpperList & SignList

    MaxNum = Len(TempChars)
    For i = 1 To MaxLength
        '#2020/10/08 同じパスワードが不定期に出現するためコメントアウト
        '原因は不明だがシード値の変更パターンにも規則がある模様
        'Randomizeステートメントは、引数を省略するとTimer関数の戻り値を取る仕様
'        Randomize
        tempPassword = Mid(TempChars, Int((MaxNum - 1 + 1) * Rnd() + 1), 1) & tempPassword
    Next
    MakePassword = tempPassword
End Function

'ZIPファイルを作成
'引数 TargetPath：圧縮するファイルまたはフォルダーのパス
'     ZipFilePath：ZIPファイルのパス
'     vPassword：パスワード　省略可
'返り値　成功したら True、失敗したらFalse
Private Function SendToZIP(ByVal TargetPath As String, ByVal ZipFilePath As String, Optional vPassword As String = "") As Boolean
    Dim vCmd As String
    '圧縮率が高いと開けないケースがある 標準モードで圧縮
    vCmd = "a -tzip -mx5 -hide "
'        vCmd = "a -tzip -mx9 -hide "
    If vPassword <> "" Then vCmd = vCmd & "-p" & vPassword & " "
    vCmd = vCmd & SetQuatation(ZipFilePath) & " " & SetQuatation(TargetPath)
    SendToZIP = (DoSevenZip(vCmd) = 0)
End Function

Private Function DoSevenZip(ByVal vCmd As String) As Long
    Dim vRet As String * 1024

    DoSevenZip = SevenZip(0&, vCmd, vRet, 1024&)
End Function

Private Function SetQuatation(ByVal Text As String) As String
    SetQuatation = """" & Replace(Text, """", """""") & """"
End Function


'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：HasErrDomain
'//  機能　　　　　　：宛先に対象のDOMAINがあるかチェック
'//  引数　　　　　　：CheckAddress：宛先、vList：対象Domainのリスト
'//  戻り値　　　　　：WhiteList外のドメインを含む：True
'//  作成日　　　　　：2016/10/03
'//  備考　　　　　　：該当がない場合は空文字
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Function HasErrDomain(ByVal CheckAddress As String, ByVal vDomainList As Variant) As String
    Dim whiteList As Variant
    Dim targetList As Variant
    Dim tempAddress As String
    Dim pos As Long
    Dim i As Long

    whiteList = Replace(vDomainList, vbCrLf, " ")
    targetList = Split(CheckAddress, ";")
    If whiteList = vbNullString Then
        For i = LBound(targetList) To UBound(targetList)
            HasErrDomain = HasErrDomain & ";" & targetList(i)
        Next
    Else
        For i = LBound(targetList) To UBound(targetList)
            pos = InStr(targetList(i), "@")
            tempAddress = Mid$(targetList(i), pos)
            If InStr(whiteList, tempAddress) = 0 Then
                HasErrDomain = HasErrDomain & ";" & targetList(i)
            End If
        Next
    End If

    If Left$(HasErrDomain, 1) = ";" Then
        HasErrDomain = Mid$(HasErrDomain, 2)
    End If
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：GetCheckDomain
'//  機能　　　　　　：チェック対象のDOMAIN一覧を返す
'//  引数　　　　　　：なし
'//  戻り値　　　　　：DOMAINの配列
'//  作成日　　　　　：2015/09/10
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Function GetCheckDomain() As Variant
    Const LIST_FILE_NAME As String = "C:\DomainList\DomainList.txt"
    Dim vFile As String
    Dim buf() As Byte

    vFile = LIST_FILE_NAME
    If FileLen(vFile) = 0 Then
        GetCheckDomain = vbNullString
    Else
        Open vFile For Binary As #1
        ReDim buf(1 To VBA.LOF(1))
        Get #1, , buf
        Close #1

        GetCheckDomain = StrConv(buf, vbUnicode, 1041&)
    End If
End Function

'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：GetNameAddrEx
'//  機能　　　　　　：受信者オブジェクトからアドレスを取得
'//  引数　　　　　　：oRecipient：対象の受信者オブジェクト
'//  戻り値　　　　　：アドレス
'//  作成日　　　　　：2014/09/17
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Function GetNameAddrEx(ByRef oRecipient As Recipient) As String
    Const ORIGINAL_DISPLAY_NAME = "http://schemas.microsoft.com/mapi/proptag/0x3a13001e"
    Const SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39fe001e"

    Dim vAddress As String
    Dim vExpanded As String

    If oRecipient.AddressEntry.Type = "SMTP" Then
        vAddress = oRecipient.Address
    ElseIf oRecipient.AddressEntry.AddressEntryUserType = olOutlookDistributionListAddressEntry Then
        vExpanded = ""
        GetNameAddrEx = ExpandGroup(oRecipient, vExpanded)
        Exit Function
    Else    ' Exchange用
        If oRecipient.AddressEntry.AddressEntryUserType = olOutlookContactAddressEntry Then
            vAddress = oRecipient.AddressEntry.PropertyAccessor.GetProperty(ORIGINAL_DISPLAY_NAME)
        Else
            vAddress = oRecipient.AddressEntry.PropertyAccessor.GetProperty(SMTP_ADDRESS)
        End If
    End If
    GetNameAddrEx = vAddress

End Function


'// ------------------------------------------------------------------------------------------------------------------------
'//  プロシージャ名　：ExpandGroup
'//  機能　　　　　　：「連絡先グループ」を展開する
'//  引数　　　　　　：oRecipient：対象の受信者オブジェクト、
'//  戻り値　　　　　：アドレス
'//  作成日　　　　　：2014/09/17
'//  備考　　　　　　：
'//  更新日：内容　　：
'// ------------------------------------------------------------------------------------------------------------------------
Private Function ExpandGroup(ByRef oRecipient As Recipient, ByRef vExpanded As String) As String
    Dim vCbLo As String
    Dim vCbHi As String
    Dim vCb As Long
    Dim vEntryID As String
    Dim vDistList As DistListItem
    Dim oMember As Recipient
    Dim vNames As String
    Dim i As Long

    If vExpanded = "" Then    ' 展開済みのグループがない = トップのグループ
        ' 65 文字目からの 4 文字がエントリー ID の長さ
        vCbLo = Mid(oRecipient.AddressEntry.ID, 65, 2)
        vCbHi = Mid(oRecipient.AddressEntry.ID, 67, 2)
        vCb = Val("&H" & vCbHi & vCbLo)

        ' 73 文字目からがアイテムのエントリー ID
        vEntryID = Mid(oRecipient.AddressEntry.ID, 73, vCb * 2)
        Set vDistList = Session.GetItemFromID(vEntryID)
    Else    ' 入れ子になっているグループの場合は 43 文字目からがアイテムのエントリー ID
        vEntryID = Mid(oRecipient.AddressEntry.ID, 43)
    End If
    '
    If InStr(vExpanded, vEntryID) > 0 Then
        Exit Function
    Else
        vExpanded = vExpanded & vEntryID & ";"
    End If
    '
    Set vDistList = Session.GetItemFromID(vEntryID)
    vNames = ""
    For i = 1 To vDistList.MemberCount
        Set oMember = vDistList.GetMember(i)
        If oMember.Address = "Unknown" Then
            vNames = vNames & ExpandGroup(oMember, vExpanded)
        Else
            vNames = vNames & GetNameAddrEx(oMember)
        End If
    Next

    ExpandGroup = vNames
End Function

'埋め込み画像ファイルかどうかチェック（通常、埋め込み画像は添付ファイル扱いになるため）
'https://social.technet.microsoft.com/Forums/exchange/en-US/542f0c0c-4de3-4752-9390-db33b8e45bca/external-email-attachments-always-seen-as-inline-attachments-with-exchange-2010?forum=exchange2010
'http://office-outlook.com/outlook-forum/index.php/t/63393/
'http://stackoverflow.com/questions/12255948/vsto-differentiate-attachments
'https://docs.aspose.com/display/emailandroid/Differentiate+between+Inline+and+Regular+Attachments

Private Function IsAttachEmbedded(ByVal oAttach As Attachment) As Boolean
    Const PR_ATTACH_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x37140003"
    Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
    
    '##Test
'    Const PR_ATTACH_CONTENT_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001E"
    Const PR_ATTACH_METHOD = "http://schemas.microsoft.com/mapi/proptag/0x37050003"

    Dim AttFlg As Long
    Dim AttCID As String
'    Dim AttLocation As String
    Dim AttMethod As Long
'
    ' 既定は通常の添付ファイル
    IsAttachEmbedded = False

    ' フラグが4（ATT_MHTML_REF）なら埋め込み画像
    AttFlg = oAttach.PropertyAccessor.GetProperty(PR_ATTACH_FLAGS)
    
    ' Content ID があれば埋め込み画像
    '「保護されたビュー」になるファイルもCONTENT IDがある
    '埋め込み画像はimageがつく戻り値になる
    AttCID = oAttach.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
    If InStr(AttCID, "image") = 0 Then
        AttCID = ""
    End If
    '#Test
'    AttLocation = oAttach.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_LOCATION)

    'RTFでがPR_ATTACH_METHOが6（ATTACH_OLE）だと埋め込み画像
    AttMethod = oAttach.PropertyAccessor.GetProperty(PR_ATTACH_METHOD)
    '#
    
    ' OLE オブジェクトなら埋め込み画像
'    If AttFlg <> 0 Or AttCID <> "" Or oAttach.Type = olOLE Then
    If AttFlg = 4 Or AttCID <> "" Or AttMethod = 6 Or oAttach.Type = olOLE Then
        IsAttachEmbedded = True
    End If
End Function

'ファイル名に使えない文字があるかチェック
Private Function CheckFileName(ByVal FileName As String) As Boolean
    Dim vExcept As Variant
    Dim i As Long

    '使用不可の文字を配列に代入
    vExcept = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(vExcept) To UBound(vExcept)
        If InStr(1, FileName, vExcept(i), vbBinaryCompare) > 0 Then
            CheckFileName = False
            Exit Function
        End If
    Next

    CheckFileName = True
End Function

Private Function SplitPath(ByVal FullPath As String, Optional ByVal ResultType As Long = 2) As String
    Dim SplitPos As Long, DotPos As Long
    SplitPos = InStrRev(FullPath, "/")
    DotPos = InStrRev(FullPath, ".")
    
    Select Case ResultType
        Case 0
            SplitPath = Left(FullPath, SplitPos - 1)
        Case 1
            If DotPos = 0 Then DotPos = Len(FullPath) + 1
            SplitPath = Mid(FullPath, SplitPos + 1, DotPos - SplitPos - 1)
        Case 2
            If DotPos = 0 Then DotPos = Len(FullPath)
            SplitPath = Mid(FullPath, DotPos + 1)
        Case Else
        Err.Raise vbObjectError + 1, "SplitPath Function", "Invalid Parameter!"
    End Select
End Function

Private Sub outputLog_test()
    OutputLog "テスト", "outputlog_test"
End Sub

'ログファイル出力
Public Sub OutputLog(ByVal msg As String, Optional ByVal ModuleName As String)
    Const LOG_FOLDER_NAME As String = "OutLookLog"
    Const LOG_FILE_NAME As String = "Log"
    Dim objFSO As Object
    Dim objLogFile As Object
    Dim TempPath As String
    Dim LogFilePath As String
    Dim objWSH As Object
    
    'ログの保存場所を確認
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objWSH = CreateObject("WScript.Shell")

    TempPath = objWSH.SpecialFolders("Desktop") & "\" & LOG_FOLDER_NAME
    ' 作業フォルダの作成 (%TEMP% で指定される一時フォルダの下に作成）
'    TempPath = objFSO.GetSpecialFolder(2) & "\" & LOG_FOLDER_NAME
    If Not objFSO.FolderExists(TempPath) Then
        objFSO.CreateFolder TempPath
    End If
    
    'ログファイルがなければ作る
    LogFilePath = TempPath & "\" & LOG_FILE_NAME & "_" & Format(Date, "yyyymmdd") & ".log"
    
    If Not objFSO.FileExists(LogFilePath) Then
        objFSO.CreateTextFile LogFilePath
    End If
    
    '追記で開く
    Set objLogFile = objFSO.OpenTextFile(LogFilePath, 8)
    '日時＋タブ＋メッセージを書き込む
    objLogFile.WriteLine Now & vbTab & ModuleName & vbTab & msg
    
    Set objLogFile = Nothing
    Set objFSO = Nothing
End Sub



