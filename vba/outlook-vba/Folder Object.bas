Sub DisplayMail()

'Declare our Variables
Dim oLookFldrInbox, oLookFldrJunk As Folder
Dim oLookName As NameSpace
Dim oLookTbl As Table
Dim oRow As Row

'Set the namespace
Set oLookName = Application.GetNamespace("MAPI")

'Define an outlook folder, in this case the inbox.
Set oLookFldrInbox = oLookName.GetDefaultFolder(olFolderInbox)

'olFolderCalendar                    9   The Calendar folder.
'olFolderConflicts                  19  The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderContacts                   10  The Contacts folder.
'olFolderDeletedItems                3   The Deleted Items folder.
'olFolderDrafts                     16  The Drafts folder.
'olFolderInbox                       6   The Inbox folder.
'olFolderJournal                    11  The Journal folder.
'olFolderJunk                       23  The Junk E-Mail folder.
'olFolderLocalFailures              21  The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderManagedEmail               29  The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
'olFolderNotes                      12  The Notes folder.
'olFolderOutbox                      4   The Outbox folder.
'olFolderSentMail                    5   The Sent Mail folder.
'olFolderServerFailures             22  The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
'olFolderSuggestedContacts          30  The Suggested Contacts folder.
'olFolderSyncIssues                 20  The Sync Issues folder. Only available for an Exchange account.
'olFolderTasks                      13  The Tasks folder.
'olFolderToDo                       28  The To Do folder.
'olPublicFoldersAllPublicFolders    18  The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
'olFolderRssFeeds                   25  The RSS Feeds folder.

'Define an outlook folder, in this case the junk.
Set oLookFldrJunk = oLookName.GetDefaultFolder(olFolderJunk)

    'Grab the path to the folder
    Debug.Print oLookFldrJunk.FolderPath

    'Is webview on?
    Debug.Print oLookFldrJunk.WebViewOn
   
    'Grab the url
    Debug.Print oLookFldrJunk.WebViewURL
    
    'How do we show item count?
    Debug.Print oLookFldrJunk.ShowItemCount
    
    'olNoItemCount           0   No item count displayed.
    'olShowTotalItemCount    2   Shows count of total number of items.
    'olShowUnreadItemCount   1   Shows count of unread items.
    
    'We could always change it if we want.
    'oLookFldrJunk.ShowItemCount = olShowTotalItemCount
    
    'How many unread emails do we have?
    Debug.Print oLookFldrJunk.UnReadItemCount
    
'Define an outlook folder, in this case the inbox.
Set oLookFldrInbox = oLookName.GetDefaultFolder(olFolderInbox)

    'Create table for our Inbox
    Set oLookTbl = oLookFldrInbox.GetTable
    
        'Let's take a look at all the columns of our table
        For Each Column In oLookTbl.Columns
            Debug.Print Column.Name
        Next
    
    'Loop through the table
    Do Until (oLookTbl.EndOfTable)
        
        'Grab the row
        Set oRow = oLookTbl.GetNextRow()
        
            'Print the details
            Debug.Print oRow("Subject")
            Debug.Print oRow("EntryID")
            Debug.Print oRow("MessageClass")
            
        'Alternative way to grab a value
        Debug.Print oRow.GetValues(2)
    
    Loop
    
End Sub
