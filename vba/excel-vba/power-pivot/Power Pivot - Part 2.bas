Sub AddNewRelationship()

'Declare Variables
Dim WrkBookConn As WorkbookConnection
Dim myModel As Model
Dim ModelRelt As ModelRelationship
Dim ModelRelts As ModelRelationships

'Create a reference to the connections in our workbook.
Set WrkBookConn = ActiveWorkbook.Connections.Item("Query - PriceData")

'Create a reference to the data model in our workbook
Set myModel = ActiveWorkbook.Model
    'myModel.AddConnection ConnectionToDataSource:=WrkBookConn
    
'Create Variables to house Model Tables.
Dim ModelTbl1 As ModelTable
Dim ModelTbl2 As ModelTable

'Get the necessary model tables.
Set ModelTbl1 = myModel.ModelTables.Item("Table1")
Set ModelTbl2 = myModel.ModelTables.Item("Table3")

'Create variables to get model table columns.
Dim PrimCol As ModelTableColumn
Dim ForgCol As ModelTableColumn

'Get the necessary model table columns.
Set ForgCol = ModelTbl1.ModelTableColumns.Item("Item")
Set PrimCol = ModelTbl2.ModelTableColumns.Item("Item")

'Add a new relationship
myModel.ModelRelationships.Add ForeignKeyColumn:=ForgCol, PrimaryKeyColumn:=PrimCol

'Create a reference to the relationships in your model.
Set ModelRelts = myModel.ModelRelationships

'Print if they're active or not. TRUE means active.
For Each ModelRelt In ModelRelts
    Debug.Print ModelRelt.Active
Next

End Sub
