Attribute VB_Name = "ImportTextFile"
Option Explicit

Sub CreatDataTable()

With ActiveSheet.QueryTables.Add(Connection:="TEXT;C:\Users\Alex\Desktop\SalesData.txt", Destination:=Range("$A$1"))

        'Name of the Query
        .Name = "SalesData"
        
        'True if field names from the data source appear as column headings for the returned data.
        .FieldNames = True
        
        'True if row numbers are added as the first column of the specified query table.
        .RowNumbers = False
        
        'True if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed.
        .FillAdjacentFormulas = False
        
        'True if any formatting common to the first five rows of data are applied to new rows of data in the query table.
        .PreserveFormatting = True
        
        'True if the query table is automatically updated each time the workbook is opened.
        .RefreshOnFileOpen = False
        
        'Returns or sets the way rows on the specified worksheet are added or deleted to accommodate the number of rows in a
        'recordset returned by a query.
        .RefreshStyle = xlInsertDeleteCells
        
        'True if password information in an ODBC connection string is saved with the specified query.
        .SavePassword = False
        
        'True if data for the QueryTable report is saved with the workbook.
        .SaveData = True
        
        'Adjust column widths on REFRESH, The maximum column width is two-thirds the width of the screen.
        .AdjustColumnWidth = True
        
        'Returns or sets the number of minutes between refreshes.
        .RefreshPeriod = 0
        
        'True if you want to specify the name of the imported text file each time the query table is refreshed.
        .TextFilePromptOnRefresh = False
        
        'Returns or sets the origin of the text file you are importing into the query table.
        'This property determines which code page is used during the data import.
        .TextFilePlatform = xlWindows
        
            'xlMacintosh 1   Macintosh
            'xlMSDOS     3   MS-DOS
            'xlWindows   2   Microsoft Windows
        
        'Returns or sets the row number at which text parsing will begin when you import a text file into a query table.
        'Valid values are integers from 1 through 32767. The default value is 1.
        .TextFileStartRow = 1
        
        'Returns or sets the column format for the data in the text file that you are importing into a query table.
        .TextFileParseType = xlDelimited 'Specify that it is delimting character not fixed width
        
        'Returns or sets the text qualifier when you import a text file into a query table.
        'The text qualifier specifies that the enclosed data is in text format.
        .TextFileTextQualifier = xlTextQualifierNone
        
        'The File Delimeter Types
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = "" & Chr(10) & ""
                
        'Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns
        'in the text file that you are importing into a query table. The default constant for each column is xlGeneral.
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1) 'All have the general type.
        
         'xlGeneralFormat General
         'xlTextFormat    Text
         'xlSkipColumn    Skip column
         'xlDMYFormat     Day-Month-Year date format
         'xlDYMFormat     Day-Year-Month date format
         'xlEMDFormat     EMD date
         'xlMDYFormat     Month-Day-Year date format
         'xlMYDFormat     Month-Year-Day date format
         'xlYDMFormat     Year-Day-Month date format
         'xlYMDFormat     Year-Month-Day date format
        
        'True for Microsoft Excel to treat numbers imported as text that begin with a "-" symbol as a negative symbol.
        .TextFileTrailingMinusNumbers = True
        
        'Refresh the query, but don't have it refresh in the background.
        .Refresh BackgroundQuery:=False
    End With
    
    
End Sub

