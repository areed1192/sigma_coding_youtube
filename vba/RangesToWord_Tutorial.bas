Attribute VB_Name = "Tutorial"
Sub CopyRangeToWord_Multi()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Variables
    Dim Rng As Variant
    'Dim ExcRng As Range
    Dim RngArray As Variant
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new document in my application
    Set WrdDoc = WrdApp.Documents.Add
    
    'Populate my range array
    RngArray = Array(Sheet1.Range("B2:E5"), Sheet1.Range("B7:E10"), Sheet2.Range("B2:E5"), Sheet2.Range("B7:E10"))
    
    'Loop through each element in the range array
    For Each Rng In RngArray
    
        'Create a reference to the range I want to Copy.
        'Set ExcRng = Rng
         Rng.Copy
            
        'Pause the Excel Application for 2 seconds
        Application.Wait Now() + #12:00:03 AM#
        
        'With the current selection paste the range
        With WrdApp.Selection
            .PasteSpecial DataType:=wdPasteOLEObject, Link:=True
        End With
        
        'Create a new page
        WrdApp.ActiveDocument.Sections.Add
        
        'Go to the newly created page
        WrdApp.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext
        
        'Clear my clipboard
        Application.CutCopyMode = False
    
    Next
        
End Sub












