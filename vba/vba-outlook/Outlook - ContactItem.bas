Option Explicit

Sub WorkingWithContactItems()

'Declare our Variables
Dim oLookItem As Object
Dim oLookName As NameSpace
Dim oLookFldr As Folder
Dim oLookContactItem As ContactItem

'Set the namespace
Set oLookName = Application.GetNamespace("MAPI")

'Define the folder that contains our contacts.
Set oLookFldr = oLookName.GetDefaultFolder(olFolderContacts)

'Loop through all the items in the folder.
For Each oLookItem In oLookFldr.Items
    
    'First Make sure it's a contact Item. The Contact Folder also has distribution list.
    If oLookItem.Class = olContact Then
        
        'Reassign it.
        Set oLookContactItem = oLookItem
            
            'Print some details.
            Debug.Print oLookContactItem.Birthday
            Debug.Print oLookContactItem.FullName
            Debug.Print oLookContactItem.MailingAddress
            Debug.Print oLookContactItem.CompanyName
            Debug.Print "-------------------------------------"
    
    End If

Next

'Display the Item.
oLookContactItem.Display

'Save it as `Doc` item. VERY IMPORTANT IT'S DOC AND NOT DOCX
oLookContactItem.SaveAs path:="C:\Users\Alex\OneDrive\Desktop\" + oLookContactItem.FirstName + ".doc", Type:=olDoc

End Sub
