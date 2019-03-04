Option Explicit

Sub CopyTableToWord_Single()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim WrdTbl As Word.Table
    
    'Declare Excel Variables
    Dim ExcLisObj As ListObject
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new document in my application
    Set WrdDoc = WrdApp.Documents.Add
      
    'Define ListObject
    Set ExcLisObj = ActiveSheet.ListObjects(1)

    'Create a reference to the range I want to Copy.
     ExcLisObj.Range.Copy
        
    'Pause the Excel Application for 2 seconds
    Application.Wait Now() + #12:00:01 AM#
    
    
    With WrdApp.Selection
        .PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
    End With
    
    Set WrdTbl = WrdDoc.Tables(WrdDoc.Tables.Count)
        WrdTbl.AllowAutoFit = True
        WrdTbl.AutoFitBehavior (wdAutoFitWindow)
        WrdTbl.Spacing = 19
        WrdTbl.Shading.BackgroundPatternColorIndex = wdBlue
        
    
    'Create a new page
    WrdApp.ActiveDocument.Sections.Add
    
    'Go to the newly created page
    WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToNext
    
    'Clear my clipboard
    Application.CutCopyMode = False

    
    WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToFirst

End Sub

Sub CopyTableToWord_Multi()

    'Declare Word Variables
    Dim WrdApp As Word.Application
    Dim WrdDoc As Word.Document
    Dim WrdTbl As Word.Table
    
    'Declare Excel Variables
    Dim ExcLisObj As ListObject
    
    'Create a new instance of Word
    Set WrdApp = New Word.Application
        WrdApp.Visible = True
        WrdApp.Activate
        
    'Create a new document in my application
    Set WrdDoc = WrdApp.Documents.Add
      
    'Loop through each element in the range array
    For Each ExcLisObj In ActiveSheet.ListObjects
    
        'Create a reference to the range I want to Copy.
         ExcLisObj.Range.Copy
            
        'Pause the Excel Application for 2 seconds
        Application.Wait Now() + #12:00:01 AM#
        
        
        With WrdApp.Selection
            .PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
        End With
        
        Set WrdTbl = WrdDoc.Tables(WrdDoc.Tables.Count)
            WrdTbl.AllowAutoFit = True
            WrdTbl.AutoFitBehavior (wdAutoFitWindow)
            WrdTbl.Spacing = 19
            WrdTbl.Shading.BackgroundPatternColorIndex = wdBlue
            
        
        'Create a new page
        WrdApp.ActiveDocument.Sections.Add
        
        'Go to the newly created page
        WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToNext
        
        'Clear my clipboard
        Application.CutCopyMode = False
    
    Next
    
    WrdApp.Selection.GoTo What:=wdGoToPage, which:=wdGoToFirst

End Sub
