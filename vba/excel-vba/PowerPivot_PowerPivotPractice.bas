Attribute VB_Name = "PowerPivotPractice"
Sub PowerPivot()

'Declare your variables.
Dim myModel As Model
Dim myModelTables As ModelTables
Dim myModelMeasures As ModelMeasures
Dim myModelTable As ModelTable
Dim wbcon As WorkbookConnection
Dim wbcons As Connections

'Create a reference to a connection
'Set wbcons = ActiveWorkbook.Connections
'Set wbcon = wbcons.Item(Index:=1)

'Create a reference to the PowerPivot Model
Set myModel = ActiveWorkbook.Model

'Add a connection to the model
'myModel.AddConnection ConnectionToDataSource:=wbcon

'Get all the tables in my PowerPivot Model
Set myModelTables = myModel.ModelTables
    
'Loop through each table and print the table name.
For Each myModelTable In myModelTables
   Debug.Print myModelTable.Name
Next

Set myModelTable = myModelTables.Item(1)
'Get all the tables in my PowerPivot Model
Set myModelMeasures = myModel.ModelMeasures
    myModelMeasures.Add "TotalTransactionCount2", myModelTable, "Sum(Table1[Transaction])", myModel.ModelFormatDecimalNumber, "This is count of all my transactions."





    

'Workbooks("ChartsToOutlook.xlsm").Connections.Add2 _
'        Name:="MyNewConnection", _
'        Description:="This is my bakery dataset.", _
'        ConnectionString:="WORKSHEET;https://petco-my.sharepoint.com/personal/305197_petco_com/Documents/Personal - YouTube Videos/Lessons - VBA/ChartsToOutlook.xlsm", _
'        CommandText:="ChartsToOutlook.xlsm!Table13", _
'        lCmdtype:=7, _
'        CreateModelConnection:=True, _
'        ImportRelationships:=False



End Sub

Sub Test()

Dim wbcon As WorkbookConnection
Dim wbcons As Connections

Set wbcons = ActiveWorkbook.Connections
'
'For Each wbcon In wbcons
'    Debug.Print wbcon.Name
'Next

Set wbcon = wbcons.Item("Query - PriceData")
Debug.Print wbcon.Name

'Set wbcon = wbcons.Item(Index:=1)
'Debug.Print wbcon.Name
'Debug.Print wbcon.Creator
''Debug.Print wbcon.DataFeedConnection
'Debug.Print wbcon.InModel
'Debug.Print wbcon.Application
''Debug.Print wbcon.Ranges
''Debug.Print wbcon.Parent
'Debug.Print wbcon.Type
'Debug.Print wbcon.Description
''Debug.Print wbcon.ModelTables
''Debug.Print wbcon.WorksheetDataConnection
''Debug.Print wbcon.TextConnection
'
''Debug.Print wbcon.OLEDBConnection
'
'
'Set wbcon = Model.AddConnection(ConnectionToDataSource:=wbcon)
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
