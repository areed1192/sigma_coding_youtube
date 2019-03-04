Attribute VB_Name = "Module1"
Option Explicit

Sub ExportChartToWord()

    'Declare Word Object Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    
    'Declare Excel Object Variables
    Dim Chrt As ChartObject
    
    'Create a New Instance Of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
    
    'Create a new Word Document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Create a Reference to the chart I want to Export
    Set Chrt = ActiveSheet.ChartObjects(1)
        Chrt.Chart.ChartArea.Copy

    'Paste into Word Document
    With WrdApp.Selection
        .PasteSpecial Link:=True, DataType:=wdPasteOLEObject
    End With

End Sub

Sub ExportingToWord_MultipleCharts_Worksheet()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim SecCnt As Integer

    'Declare Excel Variables
    Dim ChrtObj As ChartObject
    Dim Rng As Range
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Loop through the charts on the active sheet
    For Each ChrtObj In ActiveSheet.ChartObjects
    
        'Copy the chart
        ChrtObj.Chart.ChartArea.Copy

        'Paste the Chart in the Word Document
        With WrdApp.Selection
            .PasteSpecial Link:=True, DataType:=wdPasteOLEObject, Placement:=wdInLine
        End With
        
        'Count the pages in the Word Document
        SecCnt = WrdApp.ActiveDocument.Sections.Count
        
        'Add a new page to the document.
         WrdApp.ActiveDocument.Sections.Add

        'Go to the newly created page.
        WrdApp.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext
        
    Next ChrtObj

End Sub


Sub ExportingToWord_MultipleCharts_Workbook()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim SecCnt As Integer

    'Declare Excel Variables
    Dim ChrtObj As ChartObject
    Dim WrkSht As Worksheet
    Dim Rng As Range
    Dim ChrCnt As Integer
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new word document
    Set WrdDoc = WrdApp.Documents.Add
    
    ChrCnt = 0
    
    'Loop through all the worksheets in the Workbook that contains this code.
    For Each WrkSht In ThisWorkbook.Worksheets
    
        'Fix the instability error
        WrkSht.Activate
    
        'Loop through the charts on the active sheet
        For Each ChrtObj In WrkSht.ChartObjects
        
            'Copy the chart
            ChrtObj.Chart.ChartArea.Copy
            
            'Increment Chart Count
            ChrCnt = ChrCnt + 1
            
            'Fix the instability error
            Application.Wait Now + #12:00:01 AM#
    
            'Paste the Chart in the Word Document
            With WrdApp.Selection
                .PasteSpecial Link:=True, DataType:=wdPasteOLEObject, Placement:=wdInLine
            End With
            
            'Count the pages in the Word Document
            SecCnt = WrdApp.ActiveDocument.Sections.Count
            
            'Add a new page to the document.
            WrdApp.ActiveDocument.Sections.Add

            'Go to the newly created page.
            WrdApp.Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext
            
            'Fix instability Errors
            Application.CutCopyMode = False
            
        Next ChrtObj
    
    Next WrkSht

End Sub

Sub ExportChartToWord2()
    
    'Declare Word Object Variables
    Dim WrdApp As Object
    Dim WrdDoc As Object
    
    'Declare Excel Object Variables
    Dim Chrt As ChartObject
    
    'Create a New Instance Of Word
    Set WrdApp = CreateObject("Word.Application")
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new Word Document
    Set WrdDoc = WrdApp.Documents.Add
    
    'Create a Reference to the chart I want to Export
    Set Chrt = ActiveSheet.ChartObjects(1)
        Chrt.Chart.ChartArea.Copy
    
    'Paste into Word Document
    With WrdApp.Selection
        .PasteSpecial Link:=True, DataType:=wdPasteOLEObject
    End With

End Sub

