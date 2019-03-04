Attribute VB_Name = "Module2"
Option Explicit

Sub ExportChartToWord_SingleWorksheet()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Object Variable
    Dim ChrObj As ChartObject
       
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Loop through the Charts on my ACTIVE SHEET
    For Each ChrObj In ActiveSheet.ChartObjects
    
        'Copy the chart
        ChrObj.Chart.ChartArea.Copy
        
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
    
    Next ChrObj

End Sub


Sub ExportChartToWord_SingleWorkbook()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Object Variable
    Dim ChrObj As ChartObject
    Dim WrkSht As Worksheet
       
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Loop through each WorkSheet
    For Each WrkSht In ThisWorkbook.Worksheets
    
        'Activate the next worksheet
        'WrkSht.Activate
    
        'Loop through the Charts on my ACTIVE SHEET
        For Each ChrObj In WrkSht.ChartObjects
        
            'Copy the chart
            ChrObj.Chart.ChartArea.Copy
            
            'Pause the application for two seconds
            Application.Wait Now + #12:00:01 AM#
            
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
        
        Next ChrObj
        
    Next WrkSht

End Sub

Sub CreateAWordDocument()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
       
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
End Sub



