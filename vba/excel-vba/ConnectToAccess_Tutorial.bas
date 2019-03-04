Attribute VB_Name = "Tutorial"
Option Explicit

Sub ExportDataToExcel()

    'Declare Variables
    Dim ConnObj As ADODB.Connection
    Dim RecSet As ADODB.Recordset
    Dim ConnCmd As ADODB.Command
    Dim ColNames As ADODB.Fields
    Dim DataSource As String
    Dim intLoop As Integer
    
    'Define the data source
    DataSource = "C:\Users\Alex\Desktop\Financial_Data.accdb"
    
    'Create a new connection object and a new command object
    Set ConnObj = New ADODB.Connection
    Set ConnCmd = New ADODB.Command
    
    'Create a new connection
    With ConnObj
        .Provider = "Microsoft.ACE.OLEDB.12.0" 'For *ACCDB. Databases
        .ConnectionString = DataSource
        .Open
    End With
    
    'Setting the Active Connection for our command object
    ConnCmd.ActiveConnection = ConnObj
    
    'Define the Query & query type
    ConnCmd.CommandText = "SELECT * FROM ACTUALS_CAPITAL"
    ConnCmd.CommandType = adCmdText
    
    'Executing the Query & get the column names.
    Set RecSet = ConnCmd.Execute
    Set ColNames = RecSet.Fields
    
    'Populate my headers
    For intLoop = 0 To ColNames.Count - 1
        Cells(1, intLoop + 1).Value = ColNames.Item(intLoop).Name
    Next
    
    'Dump the recordset
    Range("A2").CopyFromRecordset Data:=RecSet
    
    'Close the connection
    ConnObj.Close
    
End Sub






