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
