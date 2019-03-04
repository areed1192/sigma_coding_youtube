Attribute VB_Name = "Practice"
Option Explicit

Sub RangeToWord_Single()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Object Variable
    Dim ExcRng As Range
       
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
    
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add

    'Set the Range
    Set ExcRng = ActiveSheet.Range("B2:E6")
    
        'Copy the range
        ExcRng.Copy
        
    'Pause the application for two seconds
    Application.Wait Now + #12:00:02 AM#
    
    'Paste the chart in the Word Document
    WrdDoc.Paragraphs(1).Range.PasteExcelTable LinkedToExcel:=True, WordFormatting:=False, RTF:=False
    
    'Clear Clipboard
    Application.CutCopyMode = False
    
End Sub

Sub RangeToWord_Multi()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Object Variable
    Dim Rng As Variant
    Dim ExcRng As Range
    Dim RngArray As Variant
       
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Array that houses all of our ranges that we want to export
    RngArray = Array(Worksheets("Sheet1").Range("B2:E6"), Worksheets("Sheet1").Range("B7:E10"), Worksheets("Sheet2").Range("B2:E5"))

    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Loop through the Charts on my ACTIVE SHEET
    For Each Rng In RngArray
    
        'Set the Range
        Set ExcRng = Rng
        
        'Copy the range
        Rng.Copy
        
        'Pause the application for two seconds
        Application.Wait Now + #12:00:02 AM#
        
        'Paste the chart in the Word Document
        With WrdApp.Selection
            .PasteSpecial Link:=True, DataType:=wdPasteOLEObject
        End With
        
        'Add a new page
        WrdApp.ActiveDocument.Sections.Add
        
        'Go to the newly created page
        WrdApp.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext
        
        'Clear Clipboard
        Application.CutCopyMode = False
    
    Next Rng
    
End Sub



