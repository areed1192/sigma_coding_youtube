Sub WorkingWithPublisher()

'Declare our Variables.
Dim pubApp As Application
Dim pubDoc As Document

Dim pubPages As Pages
Dim pubPage As Page

Dim pubShapes As Shapes
Dim pubShape As Shape

Dim pubTable As Table

'Grab the Application
Set pubApp = Application
    
    'Print out some details about our application.
    Debug.Print pubApp.Name
    Debug.Print pubApp.Path

'Grab the Active Document.
Set pubDoc = Application.ActiveDocument
    
'Let's the Grab the Pages Collection.
Set pubPages = pubDoc.Pages
    
    'Print out the number of pages in our document.
    Debug.Print pubPages.Count

'Grab the First and Only page in my collection.
Set pubPage = pubPages.Item(1)
    
    'Print out the Page Name.
    Debug.Print pubPage.Name
    Debug.Print pubPage.Width
    Debug.Print pubPage.Height
    
'Lets add a table to Page 1.
Set pubShape = pubPage.Shapes.AddTable(NumRows:=4, NumColumns:=3, Left:=100, Top:=100, Width:=400, Height:=300)

'Grab the Table.
Set pubTable = pubShape.Table
    
    Debug.Print pubTable.Rows.Count
    Debug.Print pubTable.Columns.Count
    
'Add another page to our document.
pubDoc.Pages.Add Count:=2, After:=1

End Sub