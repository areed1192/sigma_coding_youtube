Attribute VB_Name = "Practice"
Sub ExportingToWord_MultiplePages()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim WrdRng As Word.Range
    Dim WrdShp As Word.InlineShape
    
    'Declare Excel Variables
    Dim ChrtObj As ChartObject
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Open("C:\Users\Alex\Desktop\MyWordDoc.docx")
    
    'Loop through the charts on the active sheet
    Set ChrtObj = ActiveSheet.ChartObjects(1)
    
    'Copy the chart
    ChrtObj.Chart.ChartArea.Copy
        
    'PASTE TO A BOOKMARK
    Set WrdRng = WrdDoc.Bookmarks("MyChartPosition").Range
    
    Application.Wait Now() + #12:00:03 AM#
    
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
    End With
    
    Set WrdShp = WrdDoc.InlineShapes(1)
    
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    'PASTE TO A PARAGRAPH
    Set WrdRng = WrdDoc.Paragraphs(4).Range
        WrdRng.Collapse Direction:=wdCollapseStart
    
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
        .ParagraphFormat.SpaceAfter = 40
        .ParagraphFormat.SpaceBefore = 40
    End With
    
    Set WrdShp = WrdDoc.InlineShapes(2)
    
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With
    
    'PASTE TO SECOND PAGE IN THE SECOND PARAGRAPH
    Set WrdRng = WrdDoc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    Set WrdRng = WrdRng.Next(Unit:=wdParagraph)
        WrdRng.Collapse Direction:=wdCollapseStart
    
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
    End With
    
    Set WrdShp = WrdDoc.InlineShapes(3)
        
    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
        .Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

    'PASTE TO SECTION
    Set WrdRng = WrdDoc.Sections(2).Range.Paragraphs(2).Range
        WrdRng.Collapse Direction:=wdCollapseStart
        
    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
    End With
    
    'PASTE INSIDE WORDS
    Set WrdRng = WrdDoc.Paragraphs(1).Range.Words(10)
        WrdRng.Collapse Direction:=wdCollapseStart

    With WrdRng
        .PasteSpecial DataType:=wdPasteMetafilePicture, Placement:=wdInLine
    End With

    Set WrdShp = WrdDoc.InlineShapes(1)

    With WrdShp
        .Height = WrdApp.InchesToPoints(1.5)
        .Width = WrdApp.InchesToPoints(2.5)
    End With



    'Clear the Clipboard
    Application.CutCopyMode = False
    
    Set WrdRng = Nothing

End Sub


    
'    Set WrdShp = WrdDoc.InlineShapes(WrdDoc.InlineShapes.Count)
'
'    WrdShp.Select
'
'    With WrdApp.Selection.ShapeRange
'        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'        .RelativeVerticalPosition = wdRelativeHorizontalPositionMargin
'    End With

