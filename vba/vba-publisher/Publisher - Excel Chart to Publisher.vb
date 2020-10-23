Sub ExportExcelChartToPublisher()

'Declare our Object Variables.
Dim pubApp As Publisher.Application
Dim pubDoc As Publisher.Document
Dim pubPage As Publisher.Page
Dim pubShape As Publisher.Shape

'Declare our Excel Variables.
Dim xlBook As Workbook
Dim xlSheet As Worksheet
Dim xlChart As ChartObject

'Grab the Workbook.
Set xlBook = ThisWorkbook

'Grab the Worksheet.
Set xlSheet = xlBook.Worksheets("Objects")

'Grab the Chart.
Set xlChart = xlSheet.ChartObjects(1)

'Create or grab the instance of Publisher.
On Error Resume Next
    
    'Grab the Active Instance of Publisher if it's open.
    Set pubApp = GetObject(, "Publisher.Application")
    
    'If the application is not open it will return a 429 error.
    If Err.Number = 429 Then
    
        'Clear the Error.
        Err.Clear
        
        'Create a new instance of Publisher.
        Set pubApp = New Publisher.Application
    
    End If

'Create a new Publisher Document, THIS WILL LET YOU SEE PUBLISHER!
Set pubDoc = pubApp.Documents.Add

'Grab the First Page in the document.
Set pubPage = pubDoc.Pages(1)

'Copy the Chart.
xlChart.Chart.Copy

'Pause the Excel Application for 1 second.
Application.Wait Now + #12:00:02 AM#

'Paste it to the Page.
pubPage.Shapes.Paste

End Sub
