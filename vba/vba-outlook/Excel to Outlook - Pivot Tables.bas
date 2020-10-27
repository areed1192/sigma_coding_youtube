Sub PivotTableToOutlook()

'Declare Outlook Variables.
Dim oLookApp As Outlook.Application
Dim oLookItm As Outlook.MailItem
Dim oLookIns As Outlook.Inspector

'Declare Word Variables.
Dim oWrdDoc As Word.Document
Dim oWrdRng As Word.Range

'Declare Excel Variables.
Dim PvtTbl As PivotTable
Dim PvtRng As Range
Dim FilterRng As Range
Dim FilterMonthName As Range
Dim FilterYear As Range
Dim MonthName As String
Dim YearNumber As String
Dim PvtYearField As PivotField
Dim PvtMonthField As PivotField

'Grab the Active Outlook Application if it exists.
On Error Resume Next

'Try and Grab the Active instance.
Set oLookApp = GetObject(, "Outlook.Application")
    
    'If there is an error, create a new instance of Outlook.
    If Err.Number = 429 Then
        
        'Clear the Error.
        Err.Clear
        
        'Create the Outlook App.
        Set oLookApp = New Outlook.Application
    
    End If

'Grab the Pivot Table Object.
Set PvtTbl = ThisWorkbook.Worksheets("Pivot_Table").PivotTables("MyNewPivotTable")

'Grab the Filter Range in the Pivot Table Sheet.
Set FilterRng = ThisWorkbook.Worksheets("Pivot_Table").Range("FilterMonthName")

'Grab a "smaller" slice of the filter range.
Set FilterRng = ThisWorkbook.Worksheets("Pivot_Table").Range("E2:E4")

'Loop through each Filter Month.
For Each FilterMonthName In FilterRng

    'Grab the Month Name.
    MonthName = FilterMonthName.Value
    
    'Grab the Year Number.
    YearNumber = FilterMonthName.Offset(0, 1).Value
    
    'Clear all the Filters First.
    PvtTbl.ClearAllFilters
    
    'Grab The Year Field.
    Set PvtYearField = PvtTbl.PivotFields("Year")
        
        'Set the Filter.
        PvtYearField.CurrentPage = YearNumber
        
    'Grab The Month Name Field.
    Set PvtMonthField = PvtTbl.PivotFields("Month Name")
        
        'Set the Filter.
        PvtMonthField.CurrentPage = MonthName
    
    'Grab the Pivot Table Range.
    Set PvtRng = PvtTbl.TableRange2
    Set PvtRng = PvtTbl.TableRange1
    
    'Create a New Email.
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    
    'With the new Email.
    With oLookItm
    
        'fill out the basic Info.
        .To = "abc@xyz.com"
        .CC = "abc@xyz.com"
        .Subject = "Pivot Table Report for Month End Close"
        .Body = "Here is the Pivot Table Report for the End of the Month."
        
        'Display the email
        .Display

        'Get the active inspector
        Set oLookIns = .GetInspector

        'Get the Word Editor
        Set oWrdDoc = oLookIns.WordEditor
        
        'Copy the Range.
        PvtRng.Copy
        
        'Pause it For a second or Two.
        Application.Wait Now() + #12:00:01 AM#
        
        'Define the Range in the Email we want to Paste the Pivot Table.
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
            
            'Collapse the range.
            oWrdRng.Collapse Direction:=wdCollapseEnd
            
        'Insert a new Paragraph.
        Set oWrdRng = oWrdDoc.Paragraphs.Add
            
            'Make sure there is some space between the paragaph and the content.
            oWrdRng.InsertBreak
        
        'Paste the range as an OLE Object that is linked. You could obviously change this to be something different.
        oWrdRng.PasteSpecial DataType:=wdPasteOLEObject, Link:=False

    End With
    
    'Clear The Clipboard.
    Application.CutCopyMode = False
    
Next

End Sub