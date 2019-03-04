Attribute VB_Name = "PowerPivotSettled"
Sub ReferencePPM()
Dim myModel As Model

Set myModel = ActiveWorkbook.Model
    Debug.Print myModel.Application
    Debug.Print myModel.Name
    Debug.Print myModel.Creator
    Debug.Print myModel.DataModelConnection

End Sub



Sub AddAConnection()

WrkName = ActiveWorkbook.Name
TblName = Sheets("Price_Data").ListObjects(1).Name
FilPath = ActiveWorkbook.Path

ConnStr = "WORKSHEET;" + FilPath
CommTxt = WrkName + "!" + TblName

'"WORKSHEET;C:\Users\305197\Desktop\PowerPivot.xlsm"
Workbooks("PowerPivot.xlsm").Connections.Add2 _
                                        Name:="MyConnectionToData", _
                                        Description:="This is my price dataset.", _
                                        ConnectionString:=ConnStr, _
                                        CommandText:=CommTxt, _
                                        lCmdtype:=xlCmdExcel, _
                                        CreateModelConnection:=True, _
                                        ImportRelationships:=False

End Sub


Sub AddMeasureToPowerPivot()

'Declare your variables.
Dim myModel As Model
Dim myModelTables As ModelTables
Dim myModelMeasures As ModelMeasures
Dim myModelTable As ModelTable

'Create a reference to the PowerPivot Model
Set myModel = ActiveWorkbook.Model

'Create a reference to the ModelTables collection.
Set myModelTables = myModel.ModelTables

'Now that we have a reference to the collection, let's select one of our tables using the item Method.
Set myModelTable = myModelTables.Item(1)

'We could have also used this method.
'Set myModelTable = myModel.ModelTables.Item(1)

'Let's create a new measure in our PowerPivot Model. First we need to create a reference to
'the Model Measures collection
Set myModelMeasures = myModel.ModelMeasures

    'We can now call the add method to add a new measure.
    myModelMeasures.Add MeasureName:="TotalTransactionCount", _
                        AssociatedTable:=myModelTable, _
                        Formula:="Sum(Table1[Transaction])", _
                        FormatInformation:=myModel.ModelFormatDecimalNumber, _
                        Description:="This is count of all my transactions."

End Sub


Sub AddNewConnection()

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




