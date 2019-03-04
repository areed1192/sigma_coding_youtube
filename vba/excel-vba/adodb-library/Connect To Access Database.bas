Option Explicit

Sub ExportDataToAccess()

    Dim ConnObj As ADODB.Connection
    Dim RecSet As ADODB.Recordset
    Dim ConnCmd As ADODB.Command
    Dim ColNames As ADODB.Fields
    Dim DataSource As String
    Dim intLoop As Integer
    
    'Define the data source
    DataSource = "C:\Users\Alex\Desktop\Financial_Data.accdb"

    'Create a new connection object & a new command object
    Set ConnObj = New ADODB.Connection
    Set ConnCmd = New ADODB.Command

    'Create a new connection
    With ConnObj
        .Provider = "Microsoft.ACE.OLEDB.12.0"    'For *.ACCDB Databases
        .ConnectionString = DataSource
        .Open
    End With
    
    'This will allow the command object to use the Active Connection
    ConnCmd.ActiveConnection = ConnObj

    'Define the Query String & the Query Type.
    ConnCmd.CommandText = "SELECT * FROM ACTUALS_CAPITAL;"
    ConnCmd.CommandType = adCmdText

    'Exectue the Query & Get the column Names.
    Set RecSet = ConnCmd.Execute
    Set ColNames = RecSet.Fields
    
    'Populate the header row of the Excel Sheet.
    For intLoop = 0 To ColNames.Count - 1
        Cells(1, intLoop + 1).Value = ColNames.Item(intLoop).Name
    Next
    
    'Dump the data in the worksheet.
    Range("A2").CopyFromRecordset RecSet
    
    'Close the Connection
    ConnObj.Close

End Sub
