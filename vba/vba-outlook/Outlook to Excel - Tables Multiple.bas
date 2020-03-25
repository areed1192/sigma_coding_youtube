Option Explicit

Sub ExportMultipleEmailTablesToExcel()

'Declare our Variables
Dim oLookFolder As Folder
Dim oLookInspector As Inspector
Dim oLookMailItemArray() As Variant
Dim oLookMailObject As Variant
Dim oLookMailItem As MailItem

'Declare Word Variables.
Dim oLookWordDoc As Word.Document
Dim oLookWordTbl As Word.Table

'Declare Excel Variables.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlWrkSheet As Excel.Worksheet
Dim xlRange As Excel.Range

'Grab the Current Folder
Set oLookFolder = Application.ActiveExplorer.CurrentFolder

    'Define the array that contains the emails we want to export.
    oLookMailItemArray = Array(oLookFolder.Items("My Gmail Table - Single"), _
                               oLookFolder.Items("My Gmail Table - Multiple"), _
                               oLookFolder.Items("My Outlook Table - Single"), _
                               oLookFolder.Items("My Outlook Table - Multiple"))
                     
    'Keep going if there is an error
    On Error Resume Next
    
    'Get the Active instance of Outlook if there is one
    Set xlApp = GetObject(, "Excel.Application")
    
        'If Outlook isn't open then create a new instance of Outlook
        If Err.Number = 429 Then
        
            'Clear Error
            Err.Clear
        
            'Create a new Excel App.
            Set xlApp = New Excel.Application
            
                'Make sure it's visible.
                xlApp.Visible = True
            
            'Add a new workbook.
            Set xlBook = xlApp.Workbooks.Add
            
            'Add a new worksheet.
            Set xlWrkSheet = xlBook.Worksheets.Add
    
        End If
    
    'Set the Workbook to the Active one, if Excel is already open. THIS ASSUMES WE HAVE A WORKBOOK IN THE EXCEL APP.
    Set xlBook = xlApp.ActiveWorkbook
    
    'Set the Worksheet to the Active one, if Excel is already open. THIS ASSUMES WE HAVE A WORKSHEET IN THE WORKBOOK.
    Set xlWrkSheet = xlBook.ActiveSheet
    
    'Loop through each item in the array.
    For Each oLookMailObject In oLookMailItemArray
        
        'Set the object to the MailItem object for Intellisense.
        Set oLookMailItem = oLookMailObject
        
        'Let's grab the Active Inspector.
        Set oLookInspector = oLookMailItem.GetInspector
        
        'Grab the Word Editor object, this returns the Word Object Model.
        Set oLookWordDoc = oLookInspector.WordEditor
        
        '---------------------------------------------------------------------------
        '   Here's a trick for you, what if our signature is a table?
        '   Well, if you know the structure of the signature then what you could do is
        '   add logic to identify the table by looking at certain cells. For example, my
        '   first cell should always contain "Alex Reed" in it.
        '---------------------------------------------------------------------------
        
        For Each oLookWordTbl In oLookWordDoc.Tables
            
            'Check to see if it's the signature table.
            If Not oLookWordTbl.Cell(1, 1).Range.Text Like "*Alex Reed*" Then
            
                'Grab the Last cell in the Worksheet.
                Set xlRange = xlWrkSheet.Range("A100000").End(xlUp)
                 
                'Handle the loops that come after the first, where we need to offset.
                If xlRange.Address <> "$A$1" Then
                     
                    'Offset by two rows.
                    Set xlRange = xlRange.Offset(2, 0)
                
                End If
                
                'Copy the table.
                oLookWordTbl.Range.Copy
    
                'Paste it to the sheet.
                xlWrkSheet.Paste Destination:=xlRange
            
            End If
        
        Next
        
    Next
    
    'Set the Worksheet Column Width.
    xlWrkSheet.Columns.ColumnWidth = 20
    
    'Set the Worksheet Row Height.
    xlWrkSheet.Rows.RowHeight = 20
    
    'Set the Horizontal Alignment so it's to the Left.
    xlWrkSheet.Cells.HorizontalAlignment = xlLeft
    
    'Turn off the Gridlines.
    xlApp.ActiveWindow.DisplayGridLines = False
    
End Sub
