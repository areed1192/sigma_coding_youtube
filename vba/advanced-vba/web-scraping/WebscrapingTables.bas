Sub WebScrapeTable()

Dim IEObject As InternetExplorer

'Create a new instance of the Internet Explorer Object
Set IEObject = New InternetExplorer

    'Make sure the app is visible
    IEObject.Visible = True
    
    'Navigate to a URL we specify.
    IEObject.Navigate Url:="https://en.wikipedia.org/wiki/Star_Wars", Flags:=navOpenInNewWindow
    
    'One of the things we need to understand is that loading the page can take a while.
    'We should always wait for the page to load before continuing on to the next step.
    'This Loop will keep us waiting as long as the IEObject is in a Busy state or
    'the ReadyState does not communicate complete.
    Do While IEObject.Busy = True Or IEObject.ReadyState <> READYSTATE_COMPLETE
       
       'Wait one second, and then try again
       Application.Wait Now + TimeValue("00:00:01")
       
    Loop
    
    'Print the URL we are currently at.
    Debug.Print IEObject.LocationURL
    
    'Get the HTML document for the page
    Dim IEDocument As HTMLDocument
    Set IEDocument = IEObject.Document
    
    'Declare our HTML collection that will consist of our tables.
    Dim IETables As IHTMLElementCollection
    Dim IETable As IHTMLTable
    Dim IERow As IHTMLTableRow
    Dim IECell As IHTMLTableCell

    'Initalize an incrementer for the Row Count, this comes in handy for dumping data.
    RowCount = 1
    
    'First grab all the 'table' tags in the document, and then grap a specific item using the index method.
    Set IETable = IEDocument.getElementsByTagName("table")(1)
        
        'Determine the maximum depth of our table
        max_rows = IETable.Rows.Length
        
        'Loop through each row in the table.
        For Each IERow In IETable.Rows
    
            'Determine the maximum number of columns in each row.
            If RowCount = 1 Then
                max_cols = IERow.Cells.Length
            End If
            
            'Initalize an incrementer for the Column Count, this comes in handy for dumping data.
            ColCount = 1
            
            'Loop through each cell in the row, you can sort of consider this as a column.
            For Each IECell In IERow.Cells
            
                Debug.Print "----------"
                Debug.Print IECell.innerText
                
                'Here is the tricky part, some cells span multiple columns. If they do then we need to handle them differently.
                If IECell.colSpan > 1 Then
                    
                    'Loop the number of cells that it spans, and dump the SAME VALUE.
                    For i = 1 To Application.WorksheetFunction.Min(IECell.colSpan, max_cols)
                    
                        'If it is an element that spans multiple rows
                        If IECell.rowSpan > 1 Then
                            
                            'Then populate the number of rows, just make sure not to exceed the number of rows in the table.
                            For j = RowCount To RowCount + IECell.rowSpan - 1
                                If j <= max_rows Then
                                    ThisWorkbook.Worksheets("Table_Extract").Cells(j, ColCount).Value = IECell.innerText
                                End If
                            Next
                        
                        'If it doesn't span multiple rows.
                        Else
                            
                            'Dump the data as long as it's an empty cell.
                            If ThisWorkbook.Worksheets("Table_Extract").Cells(RowCount, ColCount).Value = "" Then
                                ThisWorkbook.Worksheets("Table_Extract").Cells(RowCount, ColCount).Value = IECell.innerText
                            End If
                        
                        End If
                        
                        'Go to the next column
                        ColCount = ColCount + 1
                    
                    Next
                
                'Otherwise treat them like normal.
                Else
                    
                    'If it is an element that spans multiple rows
                    If IECell.rowSpan > 1 Then
                        
                        'Then populate the number of rows, just make sure not to exceed the number of rows in the table.
                        For i = RowCount To RowCount + IECell.rowSpan - 1
                            If i <= max_rows Then
                                If ThisWorkbook.Worksheets("Table_Extract").Cells(i, ColCount).Value = "" Then
                                    ThisWorkbook.Worksheets("Table_Extract").Cells(i, ColCount).Value = IECell.innerText
                                Else
                                    ThisWorkbook.Worksheets("Table_Extract").Cells(i, ColCount + 1).Value = IECell.innerText
                                End If
                            End If
                        Next
                    
                    'If it is an element that does not spans multiple rows
                    Else
                        
                        'Make sure it's a blank cell, and if it is dump the data other wise go to the next column after it.
                        If ThisWorkbook.Worksheets("Table_Extract").Cells(RowCount, ColCount).Value = "" Then
                            ThisWorkbook.Worksheets("Table_Extract").Cells(RowCount, ColCount).Value = IECell.innerText
                        Else
                            ThisWorkbook.Worksheets("Table_Extract").Cells(RowCount, ColCount + 1).Value = IECell.innerText
                        End If
                    
                    End If
                    
                    'Go to the next column
                    ColCount = ColCount + 1
                    
                End If
            Next
            
            'Go to the next row.
            RowCount = RowCount + 1
            
        Next
        
End Sub
