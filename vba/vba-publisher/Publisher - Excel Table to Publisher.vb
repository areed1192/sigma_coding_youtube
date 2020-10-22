Sub ExportExcelTableToPublisher()

'Declare our Variables.
Dim pubApp As Publisher.Application
Dim pubDoc As Publisher.Document
Dim pubPage As Publisher.Page
Dim pubShape As Publisher.Shape

Dim xlBook As Workbook
Dim xlSheet As Worksheet
Dim xlTable As ListObject

'Grab the workbook.
Set xlBook = ThisWorkbook

'Grab the "Object" worksheet.
Set xlSheet = xlBook.Worksheets("Objects")

'Grab the List Object on the Sheet.
Set xlTable = xlSheet.ListObjects(1)

'Create a new instance of Publisher.
On Error Resume Next

    'Grab the Active Instance of Publisher if it's there.
    Set pubApp = GetObject(, "Publisher.Application")
    
    'If the Application is not open, then we get a 429 Error.
    If Err.Number = 429 Then
    
        'Just open the application because it's not open.
        Err.Clear
        Set pubApp = New Publisher.Application

    End If
    
'Create a new document.
Set pubDoc = pubApp.Documents.Add

'Grab the First Page in the document.
Set pubPage = pubDoc.Pages(1)

'Copy the Excel Table.
xlTable.Range.Copy

'Pause the application for 1 second.
Application.Wait Now + #12:00:01 AM#

'Paste the Table on the Page.
pubPage.Shapes.Paste

End Sub
