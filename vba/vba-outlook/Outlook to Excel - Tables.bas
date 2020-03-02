Option Explicit

Sub WorkingWithAppointmentItems()

'Declare our Variables
Dim oLookInspector As Inspector
Dim oLookMailitem As MailItem

'Declare Word Variables.
Dim oLookWordDoc As Word.Document
Dim oLookWordTbl As Word.Table

'Declare Excel Variables.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlWrkSheet As Excel.Worksheet

'Grab the mail item.
Set oLookMailitem = Application.ActiveExplorer.CurrentFolder.Items("RE: My Sample Table")

'Let's grab the Active Inspector.
Set oLookInspector = oLookMailitem.GetInspector

'Grab the Word Editor object, this returns the Word Object Model.
Set oLookWordDoc = oLookInspector.WordEditor

'Create a new Excel App.
Set xlApp = New Excel.Application

    'Make sure it's visible.
    xlApp.Visible = True

'Add a new workbook.
Set xlBook = xlApp.Workbooks.Add

'Add a new worksheet.
Set xlWrkSheet = xlBook.Worksheets.Add

'Grab the Word Table.
Set oLookWordTbl = oLookWordDoc.Tables(1)

    'Copy the table.
    oLookWordTbl.Range.Copy

    'Paste it to the sheet.
    xlWrkSheet.Paste Destination:=xlWrkSheet.Range("A1")

End Sub