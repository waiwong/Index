# Note for outlook

## 1. auto reply

Sample of source code.

<p>

```vbnet
Private Sub Application_Startup()
    Set xlItems = Session.GetDefaultFolder(olFolderInbox).Items

    'For Each myItem In xlItems
        'If (myItem.Class = olMail) Then
            'myItem.Display
            'MsgBox myItem.FullName & ": " & myItem.LastModificationTime
        '    Exit For
        'End If
    'Next

End Sub

Private Sub xlItems_ItemAdd(ByVal objItem As Object)
    Dim xlReply As MailItem
    Dim xStr As String
    If objItem.Class <> olMail Then Exit Sub

    If InStr(LCase(objItem.Subject), "auto") = 0 Then
        Exit Sub
    End If

    If InStr(LCase(objItem.SenderEmailAddress), "waiwong02@outlook.com") = 0 Then
        Exit Sub
    End If

    Set xlReply = objItem.ReplyAll
    With xlReply
         xStr = "<p>" & "Hi, Your email has been received. Thank you!" & "</p>"
         xStr = xStr & "<br />objItem.Subject: " & objItem.Subject
         xStr = xStr & "<br />objItem.Sender:" & objItem.Sender
         xStr = xStr & "<br />objItem.SenderEmailAddress:" & objItem.SenderEmailAddress
         xStr = xStr & "<br />objItem.SenderName:" & objItem.SenderName
         xStr = xStr & "<br />objItem.To:" & objItem.To
         If objItem.CC <> Null Then
            xStr = xStr & "<br />objItem.CC:" & objItem.CC
         End If

         .HTMLBody = xStr & .HTMLBody
         '.Send
         .Save
    End With
End Sub
```

</p>

ref [source](https://www.datanumen.com/blogs/auto-reply-original-email-predefined-text-via-outlook-vba/)
and [microsoft vba](https://docs.microsoft.com/en-us/office/vba/api/outlook.mailitem)

### Run in Developer Mode

If you can see the Developer tab, you are running in developer mode. Otherwise, follow these steps to run in developer mode:

1. Click the File tab.
2. Click Options.
3. Click Customize Ribbon.
4. Select the Developer check box.

## 2 Signing your own macros with SelfCert.exe

### 2.1 Locating SelfCert.exe

1. Office 365 64-bit : C:\Program Files\Microsoft Office\root\Office16
2. Simply run SelfCert.exe after locating it by one of the methods listed above. It will prompt you to name the certificate. Personally, I use my username because that is the most convenient to me but you can also give it the name of your company or whatever you want.
3. After pressing OK, you'll get a SelfCert Success message.
   ![](img/selfcert-create-digital-certificate-robert.png?raw=true)

### 2.2 Signing your code

Visual Basic buttonBack in the VBA Editor (ALT+F11) where you created the macro choose;

Tools-> Digital Signature

You'll see that the current VBA project isn't signed yet. Press the Choose button and you'll get a screen to select a certificate. Now you can choose the certificate you just created.

Choose the Digital Signature for your VBA project
![](img/selfcert-current-signature.png?raw=true)

### 2.3 Important!

Now that we've signed the code and verified that the security settings are set correctly, you must close Outlook. You'll get prompted if you want save changes to your VBA project. Choose Yes. Once Outlook is fully closed start it again.

### 2.4 Running your signed macro for the first time

Select that you'll always trust the macros or documents from this publisher and you're done.

![](img/sign-macro-trust-publisher.png?raw=true)

### 2.5 Verify your macro security level

File-> Options-> Trust Center-> Trust Center Settings-> Macro Settings-> option: Notifications for digitally signed macros, all other macros disabled

[ref link](https://www.howto-outlook.com/howto/selfcert.htm)

## 3 Reply All with Attachments VBA Macro

```vbnet
Sub ReplyWithAttachments()
    Dim rpl As Outlook.MailItem
    Dim itm As Object

    Set itm = GetCurrentItem()
    If Not itm Is Nothing Then
        Set rpl = itm.Reply
        CopyAttachments itm, rpl
        rpl.Display
    End If

    Set rpl = Nothing
    Set itm = Nothing
End Sub

Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select

    Set objApp = Nothing
End Function

Sub CopyAttachments(objSourceItem, objTargetItem)
    Dim fso as object
    set fso = CreateObject("Scripting.FileSystemObject")
    Dim fldTemp As String = fso.GetSpecialFolder(2) ' TemporaryFolder
    strPath = fldTemp.Path & "\"
    For Each objAtt In objSourceItem.Attachments
        strFile = strPath & objAtt.FileName
        objAtt.SaveAsFile strFile
        objTargetItem.Attachments.Add strFile, , , objAtt.DisplayName
        fso.DeleteFile strFile
    Next

    Set fldTemp = Nothing
    Set fso = Nothing
End Sub
```

[ref link](https://www.msoutlook.info/question/564)

## 4 Move mail from folder to one for central process

From outlook client, I add several mailbox and the following code can move all mail from set folder to one for central process such delete.

Since my mailbox have Chinese characters and vba not support Chinese, I get the folder from text file (utf-8).

```vbnet
Sub MoveToOne()
    Dim objNS As Outlook.NameSpace: Set objNS = GetNamespace("MAPI")

    Dim fldr As Folder
    Dim myItem As Object
    Dim loopFolder As MAPIFolder
    Dim myDestFolder As Outlook.Folder

'    Dim oAccount As Account
'    Dim inbx As Folder
'    For Each oAccount In Session.Accounts
'        Debug.Print vbCr & "oAccount: " & oAccount
'        Set inbx = oAccount.DeliveryStore.GetDefaultFolder(olFolderInbox)
'        Debug.Print " folder: " & inbx.Name
'    Next
'    Set inbx = Nothing

    Dim outlookFileName As String, arr As Variant
    outlookFileName = "C:\Files\outlook_folder.txt"
    '
    'The text file with following line, if no Chinese char, no need file.
    '已删除邮件
    '垃圾邮件
    '广告邮件
    '订阅邮件
    'Deleted Items
    'Deleted Messages
    'Junk
    'Junk Email
    'Trash
    '
    If Dir(outlookFileName) <> "" Then
        Dim objStream, strData
        Set objStream = CreateObject("ADODB.Stream")
        objStream.CharSet = "utf-8"
        objStream.Open
        objStream.LoadFromFile (outlookFileName)
        strData = objStream.ReadText()
        arr = Split(strData, vbNewLine)
        Set objStream = Nothing
        '
        ''not for utf-8 files
        'Dim FSO As Object, MyFile As Object
        'Set FSO = CreateObject("Scripting.FileSystemObject")
        'Set MyFile = FSO.OpenTextFile(outlookFileName, 1)
        'arr = Split(MyFile.ReadAll, vbNewLine)
        'Set MyFile = Nothing
        'Set FSO = Nothing
    Else
         arr = Split("Deleted Items,Deleted Messages", ",")
    End If

    Dim tmp As MAPIFolder
    For Each fldr In objNS.Folders
'        Debug.Print fldr.Name & " -> " & fldr.FolderPath
        If fldr.Name = "abc@abc.com" Then
            For Each loopFolder In fldr.Folders
                Debug.Print loopFolder.Name
                If loopFolder.Name = "Deleted Items" Then
                    Set myDestFolder = loopFolder
                    Debug.Print myDestFolder.Name
                    For Each myItem In myDestFolder.Items
                        If (myItem.Class = olMail) Then
                            Debug.Print "SenderEmailAddress: " & myItem.SenderEmailAddress & "  Subject: " & myItem.Subject
                        End If
                    Next
                    Exit For
                End If
            Next
        End If

        ' Shows all the names so you can replace "test"
        For Each loopFolder In fldr.Folders
            'Debug.Print loopFolder.Name
            If fldr.Name = "abc@abc.com" And loopFolder.Name = "Deleted Items" Then
                Debug.Print "dest folder, skip move"
            Else
                If UBound(Filter(arr, loopFolder.Name)) > -1 Then
                    Debug.Print vbCr & loopFolder.Name
                    For Each myItem In loopFolder.Items
                        If (myItem.Class = olMail) Then
                            Debug.Print " Folder:" & fldr.Name & "->SenderEmailAddress: " & myItem.SenderEmailAddress & "  Subject: " & myItem.Subject
                            myItem.Move myDestFolder
                            'Exit For
                        End If
                    Next
                End If
            End If
        Next
    Next

    Set objNS = Nothing
    MsgBox "Done"
End Sub
```
