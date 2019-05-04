Attribute VB_Name = "UsingNamespaceTutorial"
Option Explicit

Sub WorkNamespace()

'Declare Object Variables
Dim oLookApp As Application
Dim oLookName As NameSpace
Dim oLookAcct As Account
Dim oLookContact As ContactItem

'Grab the name space object.
Set oLookName = Application.GetNamespace(Type:="MAPI")
    
    'Print some details about our namespace
    
    'Grab the user.
    Debug.Print oLookName.CurrentUser
    
    'Grab the profile name
    Debug.Print oLookName.CurrentProfileName
    
    'Are we offline?
    Debug.Print oLookName.Offline
    
    'What class object are we?
    Debug.Print oLookName.Class '0 means application object
    
    'Grab the connection
    Debug.Print oLookName.ExchangeConnectionMode
    'olCachedConnectedFull - The account is using cached Exchange mode on a Local Area Network or a fast connection with the Exchange server.
    
    'Grab the Servername
    Debug.Print oLookName.ExchangeMailboxServerName
    
    'Grab the Server Version
    Debug.Print oLookName.ExchangeMailboxServerVersion
    '<major version>.<minor version>.<build number>.<revision>
    
    'loop through the folders in the namespace
    Dim fldr As Folder
    For Each fldr In oLookName.Folders
        Debug.Print ("----------------")
        Debug.Print fldr.Name
    Next
    
    'Let's loop through all the accounts in the NameSpace
    For Each oLookAcct In oLookName.Accounts
        'Print some details
        Debug.Print "-------------"
        Debug.Print oLookAcct.UserName
        Debug.Print oLookAcct.CurrentUser
        Debug.Print oLookAcct.DisplayName
        Debug.Print oLookAcct.AccountType '0 means Exchange Account
    Next
    
    'loop through the categories in the namespace
    Dim ctgry As Category
    For Each ctgry In oLookName.Categories
        Debug.Print ("----------------")
        Debug.Print ctgry.Name
        Debug.Print ctgry.Color
    Next
    
    'Grab a contact
    Set oLookContact = oLookName.GetDefaultFolder(olFolderContacts).Items("April")

    'Dial the contact
    oLookName.Dial oLookContact

    
End Sub
