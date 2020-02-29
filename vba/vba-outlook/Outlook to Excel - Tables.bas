Sub ExportOutlookTableToExcel()

'This assumes I'm working on an email that is opened and contains a table inside of it.
'Will only export the first table.

'Declare our Variables
Dim oLookInspector As Inspector
Dim oLookMailItem As MailItem
Dim oLookName As NameSpace
Dim oLookWordEditor As Editor

'Declare Word Variables.
Dim oLookWordDoc As Word.Document
Dim oLookWordTbl As Word.Table

'Declare Excel Variables.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet

'Let's grab the Active Inspector.
Set oLookInspector = Application.ActiveInspector

'Grab the Word Editor object, this returns the Word Object Model.
Set oLookWordDoc = oLookInspector.WordEditor

'Create a new Excel App.
Set xlApp = New Excel.Application
    
    'Make sure it's visible.
    xlApp.Visible = True
    
'Add a new workbook.
Set xlBook = xlApp.Workbooks.Add

'Add a new worksheet.
Set xlSheet = xlBook.Worksheets.Add

'Grab the Word Table.
Set oLookWordTbl = oLookWordDoc.Tables(1)

    'Copy the table.
    oLookWordTbl.Range.Copy
    
    'Paste it to the sheet.
    xlSheet.Paste

End Sub