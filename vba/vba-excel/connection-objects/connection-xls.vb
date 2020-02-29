Option Explicit

Sub AddXLSConnection()

'Declare Object Variables.
Dim xlWorkbook As Workbook
Dim xlWorkbookConnections As Connections
Dim xlWorkbookConnection As WorkbookConnection

'Define a bunch of strings used for the connection object.
Dim connection_name As String
Dim connection_desc As String
Dim conn_str_type  As String
Dim conn_str_provider As String
Dim conn_str_datasource As String
Dim conn_extended_properties
Dim connection_string As String
Dim sql_query As String

'Grab the workbook
Set xlWorkbook = ThisWorkbook

'Grab the Workbook Connections Collection.
Set xlWorkbookConnections = xlWorkbook.Connections

'Readable properties.
'--------------------

'Every Connection has a few things we can define about. In this portion, let's define the connection name.
connection_name = "My Excel Workbook Connection For Sales"

'Then the description.
connection_desc = "A connection to my Excel workbook that contains all my Employee and Sales Data."


'Connection Components.
'----------------------

'Define the type of connection.
conn_str_type = "OLEDB"

'Define the Provider.
conn_str_provider = "Microsoft.ACE.OLEDB.12.0"

'Define the data source, in this case it's the path to the file.
conn_str_datasource = "C:\Users\Alex\OneDrive\Desktop\excel_workbook_datasource.xlsx"

'Define the extended properties.
conn_extended_properties = "'Excel 12.0 Macro;HDR=YES'" 'ITS VERY IMPORTANT YOU WRAP IT IN SINGLE QUOTES!

'Build the full connection string.
connection_string = conn_str_type + ";Provider=" + conn_str_provider + _
                                    ";Data Source=" + conn_str_datasource + _
                                    ";Extended Properties=" + conn_extended_properties + ";"

'Print out the connection string.
Debug.Print "Here is my full connection string:"
Debug.Print connection_string
Debug.Print "-------------"

'Define the query, for Excel Workbooks you can query entire Sheets or specific ranges on a sheet.
'In this case, query the "Sales_Data" sheet.
sql_query = "SELECT * FROM [Sales$]"

'Let's put it all together and add a new connection to our workbook.
xlWorkbookConnections.Add2 Name:=connection_name, _
                           Description:=connection_desc, _
                           ConnectionString:=connection_string, _
                           CommandText:=sql_query, _
                           lCmdType:=xlCmdSql, _
                           CreateModelConnection:=False, _
                           ImportRelationships:=False
                 
End Sub

Sub QueryXLSConnection()

'Declare Object Variables.
Dim xlWorkbook As Workbook
Dim xlWorksheet As Worksheet
Dim xlXLSTable As ListObject
Dim xlXLSQueryTable As QueryTable
Dim xlWorkbookConnections As Connections
Dim xlWorkbookConnection As WorkbookConnection
Dim xlOLEDBConnection As OLEDBConnection

'Grab the workbook
Set xlWorkbook = ThisWorkbook

'Grab the Workbook Connections Collection.
Set xlWorkbookConnections = xlWorkbook.Connections

'Set the connection using the Key Method.
Set xlWorkbookConnection = xlWorkbookConnections.Item("My Excel Workbook Connection For Sales")

'Now at this point we have a connection, but it's a little ambiguous. We need to define the type of connection object it is.
'In this case, we are lucky cause we built it and we know the connection type is OLEDB so that means we have an OLEDB Connection.
Set xlOLEDBConnection = xlWorkbookConnection.OLEDBConnection

'With my connection now selected, let's begin the process of querying it. First I need to define which sheet I want it on.
Set xlWorksheet = ThisWorkbook.Worksheets("My_Excel_File")
    
'Pull it all together.
Set xlXLSTable = xlWorksheet.ListObjects.Add(SourceType:=xlSrcExternal, _
                                             Source:=xlWorkbookConnection, _
                                             Destination:=xlWorksheet.Range("A1"))

'Great we have a list object, let's grab it's QueryTable. A QueryTable represents a worksheet table
'built from data returned from an external data source, such as a SQL server or a Microsoft Access database.
Set xlXLSQueryTable = xlXLSTable.QueryTable

'With our query table.
With xlXLSQueryTable
    
    'Define the Command Type, just use the exisiting connection object.
    .CommandType = xlOLEDBConnection.CommandType
    
    'Define the Command Text, just use the exisiting connection object.
    .CommandText = xlOLEDBConnection.CommandText
    
    'Change the display name of the table.
    .ListObject.DisplayName = "ExcelWorkbookData"

    'Refresh the data, this pulls it in. However, I don't want the query to run in the background, so set BackgroundQuery to false.
    .Refresh BackgroundQuery:=False
    
    'I don't want the columns adjusted everytime the data is pulled in.
    .AdjustColumnWidth = False
    
    'Change the style as Well.
    .ListObject.TableStyle = "TableStyleLight21"
    
    'Set the Refresh Style.
    .RefreshStyle = xlInsertDeleteCells

    'Set all the rows to a height of 25.
    .ListObject.Range.Rows.RowHeight = 25
    
    'Set all the columns to a width of 20.
    .ListObject.DataBodyRange.Columns.ColumnWidth = 20
    
End With

End Sub