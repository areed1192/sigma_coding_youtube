Sub ModelTableObject()

'Declare our variables
Dim myModel As Model
Dim ModelTbls As ModelTables
Dim ModelTbl As ModelTable
Dim ModelCols As ModelTableColumns
Dim ModelCol As ModelTableColumn
Dim ModelConn As ModelConnection

'Create a reference to our power pivot model
Set myModel = ActiveWorkbook.Model

'Reference the model tables collection
Set ModelTbls = myModel.ModelTables

    'Count the number of tables
    Debug.Print ModelTbls.Count
    
    'Get the parent object name
    Debug.Print ModelTbls.Parent.Name
    
'Loop through each table
For Each ModelTbl In ModelTbls
    Debug.Print ModelTbl.Name
    Debug.Print ModelTbl.RecordCount
    Debug.Print ModelTbl.SourceName
    Debug.Print ModelTbl.SourceWorkbookConnection
Next

Set ModelTbl = ModelTbls.Item("Price_Data")
Set ModelCols = ModelTbl.ModelTableColumns

'Looping through each col
For Each ModelCol In ModelCols
    Debug.Print ModelCol.Name
    Debug.Print ModelCol.DataType
    Debug.Print ModelCol.Application
Next

'Lets work with the model connection object now.
Set ModelConn = myModel.DataModelConnection.ModelConnection
    Debug.Print ModelConn.CommandType
    Debug.Print ModelConn.CommandText
    Debug.Print ModelConn.ADOConnection
    
End Sub
