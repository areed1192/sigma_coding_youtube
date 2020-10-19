Sub WorkingWithPages()

Dim pubApp As Application
Dim pubDoc As Document

Dim pubPage As Page
Dim pubPages As Pages

Dim pubPageBackground As PageBackground
Dim pubPageFill As FillFormat

Dim pubShape As Shape
Dim pubShapes As Shapes

Dim pubTable As Table

'Grab the Application.
Set pubApp = Application

'Grab the Document.
Set pubDoc = pubApp.ActiveDocument

'Grab the First Page.
Set pubPage = pubDoc.Pages(1)

    'Specifies the Width of the Page.
    Debug.Print "Page Width: " + CStr(pubPage.Width)
    
    'Specifies the Height of the Page.
    Debug.Print "Page Height: " + CStr(pubPage.Height)
    
    'True if the specified Page object is a trailing page of a two-page spread. Read-only Boolean.
    Debug.Print "Is Trailing: " + CStr(pubPage.IsTrailing)
    
    'True if the specified Page object is a leading page of a two-page spread. Read-only Boolean.
    Debug.Print "Is Leading: " + CStr(pubPage.IsLeading)
    
    'True if the specified page is a Microsoft Publisher wizard page. Read-only Boolean.
    'Wizard pages are special page types for certain types of Publisher wizards
    '(such as Newsletters, Catalogs, and Web Wizards) that can be inserted into
    'a publication.
    Debug.Print "Is WizardPage: " + CStr(pubPage.IsWizardPage)
    
'Grab the Page Background Object.
Set pubPageBackground = pubPage.Background

'Check to see if it Exists, if it doesn't then create it.
If pubPageBackground.Exists = False Then
    pubPageBackground.Create
End If

'Grab the Fill Format Object.
Set pubPageFill = pubPageBackground.Fill

'Lets modify the Fill Format.
With pubPageFill

    'Change the Backcolor.
    .BackColor.RGB = RGB(Red:=0, Green:=155, Blue:=99)

    'Change the Forecolor.
    .ForeColor.RGB = RGB(Red:=155, Green:=234, Blue:=0)

    'Create a Two Color Gradient.
    .TwoColorGradient Style:=msoGradientDiagonalDown, Variant:=4

End With

'Add Pages to our Document.
pubDoc.Pages.Add Count:=2, After:=1

'Delete the Second Page.
pubDoc.Pages(2).Delete

'Duplicate the "New" second Page.
pubDoc.Pages(2).Duplicate

'Print the number of Pages in the Collection.
Debug.Print "The number of Pages in the document are: " + CStr(pubDoc.Pages.Count)

'Print the Page ID.
Debug.Print "The Page ID is: " + CStr(pubPage.PageID)

'Print the Page Index.
Debug.Print "The Page ID is: " + CStr(pubPage.PageIndex)

'Print the Page Name.
Debug.Print "The Page Name is: " + CStr(pubPage.Name)


'Add some tags to our page.
pubPage.Tags.Add Name:="Creator", Value:="Alex"

'Save the document as a picture.
pubPage.SaveAsPicture Filename:="C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Lessons - VBA\VBA - Publisher\PhotoTemplate.jpg", _
                      pbResolution:=pbPictureResolutionCommercialPrint_300dpi
                      
'We can also export the document as an HTML file.
pubPage.ExportEmailHTML Filename:="C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Lessons - VBA\VBA - Publisher\EmailTemplate.html"

End Sub