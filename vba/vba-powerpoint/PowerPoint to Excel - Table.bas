Sub ExportMultiplePowerPointTablesToExcel()

'Declare our Variables
Dim PPTPres As Presentation
Dim PPTSlide As Slide
Dim PPTShape As Shape
Dim PPTTable As Table

'Declare Excel Variables.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlWrkSheet As Excel.Worksheet
Dim xlRange As Excel.Range

'Grab the Currrent Presentation.
Set PPTPres = Application.ActivePresentation
                     
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
    Set xlBook = xlApp.Workbooks("ExportFromPowerPointToExcel.xlsm")
    
    'Set the Worksheet to the Active one, if Excel is already open. THIS ASSUMES WE HAVE A WORKSHEET IN THE WORKBOOK.
    Set xlWrkSheet = xlBook.Worksheets("Table_Export")
    
    'Loop through each Slide in the Presentation.
    For Each PPTSlide In PPTPres.Slides
    
        'Loop through each Shape in Slide.
        For Each PPTShape In PPTSlide.Shapes
            
            'If the Shape is a Table.
            If PPTShape.Type = msoTable Then
                
                'Grab the Table Object, using the Table Property.
                Set PPTTable = PPTShape.Table
                    
                    'Copy the Shape.
                    PPTShape.Copy
                
                'Grab the Last cell in the Worksheet.
                Set xlRange = xlWrkSheet.Range("A100000").End(xlUp)

                'Handle the loops that come after the first, where we need to offset.
                If xlRange.Address <> "$A$1" Then

                    'Offset by two rows.
                    Set xlRange = xlRange.Offset(2, 0)

                End If

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