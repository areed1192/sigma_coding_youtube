Attribute VB_Name = "Tutorial"

Sub PositionObjectsInWord()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim WrdRng As Word.Range
    Dim WrdShp As Word.InlineShape
    
    'Declare Excel Variables
    Dim ChrObj As ChartObject
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Open the Word Document
    Set WrdDoc = WrdApp.Documents.Open("C:\Users\Alex\Desktop\MyWordDoc.docx")
    
    'Create a reference to the chart
    Set ChrObj = ActiveSheet.ChartObjects(1)
        ChrObj.Chart.ChartArea.Copy
        
    'Pausing the application For Two Seconds
    Application.Wait Now() + #12:00:02 AM#
    
    'Paste to Bookmark
    Set WrdRng = WrdDoc.Bookmarks("MyChartPosition").Range
    
    'Paste the Object
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
    End With
    
    'Create a reference to the shape
    Set WrdShp = WrdDoc.InlineShapes(1)
    
    'Set the dimensions
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    
    
    
    
    
    'Paste to Bookmark
    Set WrdRng = WrdDoc.Paragraphs(4).Range
        WrdRng.Collapse Direction:=wdCollapseStart
    
    'Paste the Object
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
        .ParagraphFormat.SpaceAfter = 40
        .ParagraphFormat.SpaceBefore = 40
    End With
    
    'Create a reference to the shape
    Set WrdShp = WrdDoc.InlineShapes(2)
    
    'Set the dimensions
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    
    
    
    
    'Paste to Page
    Set WrdRng = WrdDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    Set WrdRng = WrdRng.Next(Unit:=wdParagraph)
        WrdRng.Collapse Direction:=wdCollapseStart
    
    'Paste the Object
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
        .ParagraphFormat.SpaceAfter = 40
        .ParagraphFormat.SpaceBefore = 40
    End With
    
    'Create a reference to the shape
    Set WrdShp = WrdDoc.InlineShapes(3)
    
    'Set the dimensions
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    
    
    
    
    'Paste to Section
    Set WrdRng = WrdDoc.Sections(2).Range.Paragraphs(2).Range
        WrdRng.Collapse Direction:=wdCollapseStart
    
    'Paste the Object
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
        .ParagraphFormat.SpaceAfter = 40
        .ParagraphFormat.SpaceBefore = 40
    End With
    
    'Create a reference to the shape
    Set WrdShp = WrdDoc.InlineShapes(4)
    
    'Set the dimensions
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    
    
    
    
    
    'Paste to Word
    Set WrdRng = WrdDoc.Paragraphs(1).Range.Words(10)
        WrdRng.Collapse Direction:=wdCollapseStart
    
    'Paste the Object
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
        .ParagraphFormat.SpaceAfter = 40
        .ParagraphFormat.SpaceBefore = 40
    End With
    
    'Create a reference to the shape
    Set WrdShp = WrdDoc.InlineShapes(1)
    
    'Set the dimensions
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
End Sub












