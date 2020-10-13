Sub WorkingWithTheDocumentObject()

'Declare our variables.
Dim pubApp As Application
Dim pubDocs As Documents
Dim pubDoc As Document

'Grab the application.
Set pubApp = Application

'Grab the Documents collection.
Set pubDocs = pubApp.Documents

'Grab the Document that houses the code.
Set pubDoc = ThisDocument

    'Print the name of the document.
    Debug.Print "The document name is: " + pubDoc.Name
    
    'Print the Full path to the document.
    Debug.Print "The Document Path is: " + pubDoc.FullName
    
    'You can update OLEObjects inside of the document.
    pubDoc.UpdateOLEObjects
    
    'Grab the Active Window and bring it to the front.
    'pubDoc.ActiveWindow.Activate
    
    'Grab the Document Direction.
    Debug.Print "The Document Direction is: " + CStr(pubDoc.DocumentDirection)
    
    'Turn off the Guides.
    pubDoc.LayoutGuides.Rows = 3
    pubDoc.LayoutGuides.Columns = 3
    
    'See if the document is saved.
    Debug.Print "The Document Saved Status is: " + CStr(pubDoc.Saved)
    
    'Print the number of Redo Actions Available.
    Debug.Print "The number of ReDo Actions Available is: " + CStr(pubDoc.RedoActionsAvailable)
    Debug.Print "The number of UnDo Actions Available is: " + CStr(pubDoc.UndoActionsAvailable)
    
    'Check to see if a data source is connected to the document.
    Debug.Print "The document Data Source Connected Status is: " + CStr(pubDoc.IsDataSourceConnected)
    
    'Print the Save Format of the Document.
    Debug.Print "The Document Save Format is: " + CStr(pubDoc.SaveFormat)
    
    'Export the document as a PDF File.
    'pubDoc.ExportAsFixedFormat Format:=pbFixedFormatTypePDF, Filename:="C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Lessons - VBA\VBA - Publisher\TestWorkingWithDocumentsV2.pdf"
    
    'Do a webpage Preview.
    'pubDoc.WebPagePreview
    
    'Create a Custom UndoAction that adds a textbox to the document.
    pubDoc.BeginCustomUndoAction ActionName:="Add Textbox with Segoe Font"
        
        'Grab Page 1.
        With pubDoc.Pages(1)
        
            'Add a Rectangle.
            Set RectangleShape = .Shapes.AddShape(Type:=msoShapeRectangle, Left:=75, Top:=75, Width:=190, Height:=130)
            
            'Change the Font in the Rectangle.
            With RectangleShape.TextFrame.TextRange
            
                .Font.Size = 14
                .Font.Bold = True
                .Font.Name = "Segoe UI"
                .Text = "This is a Segoe UI Font."
            
            End With
        
        End With
    
    pubDoc.EndCustomUndoAction
    
End Sub