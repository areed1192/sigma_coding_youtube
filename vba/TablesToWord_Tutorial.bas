Attribute VB_Name = "Tutorial"

Sub ListObjectToWord_Multi()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim WrdTbl As Word.Table
    
    'Declare Excel Variables
    Dim ExcLisObj As ListObject
    Dim WrkSht As Worksheet
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new document in the application
    Set WrdDoc = WrdApp.Documents.Add
    
    'Loop through all the Worksheets in the ACTIVE WORKBOOK
    For Each WrkSht In ThisWorkbook.Worksheets
    
        'Loop through the List objects on the ACTVIESHEET
        For Each ExcLisObj In WrkSht.ListObjects
        
            'Copy the list object
            ExcLisObj.Range.Copy
            
            'Pause the Excel application for one second
            Application.Wait Now() + #12:00:03 AM#
            
            'Paste List Object into the Word Document
            With WrdApp.Selection
                .PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
            End With
            
            'Format the table
            Set WrdTbl = WrdDoc.Tables(WrdDoc.Tables.Count)
                WrdTbl.AllowAutoFit = True
                WrdTbl.AutoFitBehavior (wdAutoFitWindow)
                
            'Create a new page
            WrdApp.ActiveDocument.Sections.Add
            
            'Go to the new page
            WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToNext
            
            'Clear my clipboard
            Application.CutCopyMode = False
            
        Next
    
    Next
    
    'Go to the first page
    WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToFirst
    
End Sub







