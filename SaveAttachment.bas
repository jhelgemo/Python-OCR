Attribute VB_Name = "SaveAttachment"
Option Explicit

Sub Save_Attachment()
'Arg 1 = Folder name of folder inside your Inbox
'Arg 2 = Subfolder
'Arg 3 = File extension, "" is every file
'Arg 4 = Save folder,

    SaveEmailAttachmentsToFolder "", "", "", ""
    
End Sub

Sub SaveEmailAttachmentsToFolder(InboxFolder As String, _
                                 SubSubFolder As String, ExtString As String, DestFolder As String)
    Dim ns As NameSpace
    Dim Inbox As Folder
    Dim SubFolder As Folder
    Dim Item As Object
    Dim AttachMnt As Attachment
    Dim FileName As String
    Dim MyDocPath As String
    Dim I As Integer
    Dim wShell As Object
    Dim fs As Object
    Dim destMailFolder As Folder
    Dim stepCounter As Integer
    

    On Error GoTo errorHandler

    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    
    Set SubFolder = Inbox.Folders(InboxFolder)
    Set SubFolder = SubFolder.Folders(SubSubFolder)
    
    ' Define inboxfolder and subfolder for processed emails
    Set destMailFolder = Inbox.Folders(InboxFolder)
    Set destMailFolder = destMailFolder.Folders("")

    I = 0
    
    ' Check subfolder for messages and exit of none found
    If SubFolder.Items.Count = 0 Then
        MsgBox "There are no messages in this folder : " & InboxFolder, _
               vbInformation, "Nothing Found"
        Set SubFolder = Nothing
        Set Inbox = Nothing
        Set ns = Nothing
        Set destMailFolder = Nothing
        Exit Sub
    End If

    'Create DestFolder if DestFolder = ""
    If DestFolder = "" Then
        Set wShell = CreateObject("WScript.Shell")
        Set fs = CreateObject("Scripting.FileSystemObject")
        MyDocPath = wShell.SpecialFolders.Item("mydocuments")
        DestFolder = MyDocPath & "\" & Format(Now, "mmm-yyyy")
        If Not fs.FolderExists(DestFolder) Then
            fs.CreateFolder DestFolder
        End If
    End If
    
    'check if destfolder ends with "\"
    
    If Right(DestFolder, 1) <> "\" Then
        DestFolder = DestFolder & "\"
    End If

    ' Check each message for attachments and extensions
    

    
    
     For stepCounter = SubFolder.Items.Count To 1 Step -1
     
     Set Item = SubFolder.Items(stepCounter)
        For Each AttachMnt In Item.Attachments
            'check if attachment has defined file extention
            If LCase(Right(AttachMnt.FileName, Len(ExtString))) = LCase(ExtString) Then
                FileName = DestFolder & "_" & Format(Item.ReceivedTime, "yyyy-mmm-dd-hh-mm-ss") & AttachMnt.FileName
                AttachMnt.SaveAsFile FileName
                I = I + 1
            End If
        Next AttachMnt
        Item.Move destMailFolder
    Next stepCounter

    ' Show this message when Finished
    If I > 0 Then
        MsgBox I & " Files processed. You can find the files here : " _
             & DestFolder, vbInformation, "Finished!"
    Else
        MsgBox "No attached files in your mail.", vbInformation, "Finished!"
    End If

    ' Clear memory
ThisMacro_exit:
    Set SubFolder = Nothing
    Set Inbox = Nothing
    Set ns = Nothing
    Set fs = Nothing
    Set wShell = Nothing
    Set destMailFolder = Nothing
    Exit Sub

    ' Error information
errorHandler:
    MsgBox "An unexpected error has occurred." _
         & vbCrLf & "Please note and report the following information." _
         & vbCrLf & "Macro Name: SaveEmailAttachmentsToFolder" _
         & vbCrLf & "Error Number: " & Err.Number _
         & vbCrLf & "Error Description: " & Err.Description _
         , vbCritical, "Error!"
    Resume ThisMacro_exit

End Sub
