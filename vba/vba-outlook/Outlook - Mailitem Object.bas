Sub DisplayMail()

'Declare our Variables
Dim oLookItem As Object
Dim oLookMail As MailItem
Dim oLookFldr As Folder
Dim oLookName As NameSpace

'Set the namespace
Set oLookName = Application.GetNamespace("MAPI")

'Define the folder that contains my emails.
Set oLookFldr = oLookName.GetDefaultFolder(olFolderInbox)

    'Display the folder
    oLookFldr.Display

'Loop through all the items in the folder.
For Each itm In oLookFldr.Items

   ' If the item is a mail item
   If TypeOf itm Is MailItem Then
      Debug.Print itm.Subject
   End If

Next


'Let's work an email
Set oLookMail = oLookFldr.Items(5)

    'Grab the body format
    Debug.Print oLookMail.BodyFormat
    
    'olFormatHTML         2   HTML format
    'olFormatPlain        1   Plain format
    'olFormatRichText     3   Rich text format
    'olFormatUnspecified  0   Unspecified format

    'What time was it recieved?
    Debug.Print oLookMail.ReceivedTime

    'Who sent it?
    Debug.Print oLookMail.Sender
    
    'What was their email address?
    Debug.Print oLookMail.SenderEmailAddress
    
    'Is it unread?
    Debug.Print oLookMail.UnRead
    
    'When was it sent?
    Debug.Print oLookMail.SentOn
    
    'Print the email body text
    Debug.Print oLookMail.Body

    
    
'Define a Recipient object variable
Dim rec As Recipient
         
'Loop through all the recipients on the email
For Each rec In oLookMail.Recipients
    
    'Print their email address
    Debug.Print rec.Address
    
    'Print their name.
    Debug.Print rec.Name
    
    'Can you send them stuff?
    Debug.Print rec.Sendable
    
Next

'Define an attachment object variable
Dim oLookAtt As Attachment

'Grab an attachment item
Set oLookAtt = oLookMail.Attachments.Item(1)

    'Grab a file name
    Debug.Print oLookAtt.FileName
    
    'Check class type - 5 means attachment
    Debug.Print oLookAtt.Class
    
    'Grab the display name
    Debug.Print oLookAtt.DisplayName
    
    'Get the size of the file in bytes
    Debug.Print oLookAtt.Size
    
    'Get the file type
    Debug.Print oLookAtt.Type
    
    'olByReference   4   This value is no longer supported since Microsoft Outlook 2007. Use olByValue to attach a copy of a file in the file system.
    'olByValue       1   The attachment is a copy of the original file and can be accessed even if the original file is removed.
    'olEmbeddeditem  5   The attachment is an Outlook message format file (.msg) and is a copy of the original message.
    'olOLE           6   The attachment is an OLE document.

'Display the email
oLookMail.Display

End Sub
