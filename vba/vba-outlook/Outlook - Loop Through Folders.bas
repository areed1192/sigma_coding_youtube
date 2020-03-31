Option Explicit

Sub SaveAllAttachmentsRecursive()

'Declare our variables
Dim oLookNamespace As NameSpace
Dim oLookFolder As Folder
Dim FileSysObj As New FileSystemObject
Dim oLookSaveFolderPath As String

'Define the Save folder path.
oLookSaveFolderPath = Environ("HOMEPATH") + "\OneDrive\Desktop\outlook_back"

'Check if the folder exists, if it doesn't create it.
If Not FileSysObj.FolderExists(FolderSpec:=oLookSaveFolderPath) Then
    
    'Create the folder
    FileSysObj.CreateFolder path:=oLookSaveFolderPath
    
End If

'Grab the namespace
Set oLookNamespace = Application.GetNamespace("MAPI")

'Loop through each folder in the Namespace
For Each oLookFolder In oLookNamespace.Folders
    
    'Call the ProcessCurrentFolder SubRoutine
    Call ProcessCurrentFolder(oLookFolder)
    
Next

End Sub

Sub ProcessCurrentFolder(ByVal ParentFolder As Folder)

'Declare our variables
Dim SubFolder As Folder
Dim oLookFldrItem As Object
Dim oLookMailItem As MailItem
Dim oLookAttachment As Attachment

Dim AttDisName As String
Dim oLookSaveFolderPath As String

'Define the Save folder path.
oLookSaveFolderPath = Environ("HOMEPATH") + "\OneDrive\Desktop\outlook_back\"

'Loop through each item in the folder
For Each oLookFldrItem In ParentFolder.Items

    'Check to see if the item we are currently on in the loop is a MAILITEM!
    If TypeOf oLookFldrItem Is Outlook.MailItem Then
    
        'Re Assignment mainly for Intellisense
        Set oLookMailItem = oLookFldrItem
            
        'Check to see if there are any attachments in the email, if so continue.
        If oLookMailItem.Attachments.Count > 0 Then
            
            'Loop through each attachment in the email.
            For Each oLookAttachment In oLookMailItem.Attachments
            
                'Print the attachment Type.
                Debug.Print "Attachment Type is: " + CStr(oLookAttachment.Type)
                
                'If it's an attached file continue
                If oLookAttachment.Type = 1 Then
                
                    'Print out some info about the Attachment Object
                    Debug.Print "Attachment File Name: " + oLookAttachment.DisplayName
                    Debug.Print "Attachment MailItem Name: " + oLookMailItem.Subject
                    Debug.Print "Parent Folder Name: " + ParentFolder.Name
                    Debug.Print "Attachment Count: " + CStr(oLookMailItem.Attachments.Count)
                    
                    'Grab the Display Name.
                    AttDisName = oLookAttachment.DisplayName
                    
                    'Filter the attachments so that we only save the PowerPoint and Excel Documents.
                    If InStr(AttDisName, "xls") > 0 Or InStr(AttDisName, "ppt") > 0 Then
                    
                        'Save to the Desktop
                        oLookAttachment.SaveAsFile path:=oLookSaveFolderPath + AttDisName
                    
                    End If
                End If
            Next
        End If
    End If
Next

'Process the SubFolders Recursively
If (ParentFolder.Folders.Count > 0) Then

    'Loop through each of the SubFolders in the ParentFolder
    For Each SubFolder In ParentFolder.Folders
        
        'Call our ProcessCurrentFolder Routine Again.
        Call ProcessCurrentFolder(SubFolder)
    
    Next
End If

End Sub