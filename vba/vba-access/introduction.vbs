Option Compare Database

Sub IntroductionToAccess()

'Declare our variables.
Dim accessApp As Application
Dim accessDatabase As Database
Dim accessTable As TableDef
Dim accessRecord As Recordset

'Grab the access application.
Set accessApp = Application

'Grab the Current Database in our application.
Set accessDatabase = accessApp.CurrentDb
Set accessDatabase = Application.DBEngine.Workspaces(0).Databases(0)

Debug.Print "The Database Full Path is:" + accessDatabase.Name
Debug.Print "The Number of tables in the database are: " + CStr(accessDatabase.TableDefs.Count)

accessDatabase.TableDefs.Delete Name:="StockPrices"

'Let's create a new table definition.
Set accessTable = accessDatabase.CreateTableDef(Name:="StockPrices")

'Take the new table and add some fields to it.
With accessTable

    'Add a field for the Date.
    .Fields.Append .CreateField("date", dbDate)
    
    'Add a field for the open, close, high, low note that these are doubles.
    .Fields.Append .CreateField("open", dbDouble)
    .Fields.Append .CreateField("close", dbDouble)
    .Fields.Append .CreateField("high", dbDouble)
    .Fields.Append .CreateField("low", dbDouble)
    
End With

'Add the new table to the Table Definitions Collection.
accessDatabase.TableDefs.Append Object:=accessTable

'Refresh the Table Definitions Collection database,
accessDatabase.TableDefs.Refresh

'Once we add the table, we can add records to it. Let's open a new recordset object.
Set accessRecord = accessDatabase.OpenRecordset(Name:="StockPrices")

'Use the `addNew` method to add a new record.
accessRecord.AddNew

'Specify the different fields we defined up above.
accessRecord("date").Value = "12-26-2020"
accessRecord("open").Value = 12
accessRecord("close").Value = 12.2
accessRecord("high").Value = 12.3
accessRecord("low").Value = 11.99

'Update the Recordset.
accessRecord.Update

End Sub