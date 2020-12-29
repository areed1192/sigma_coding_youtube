Option Compare Database

Sub WorkWithQueries()

'Declare our variables.
Dim AccessApp As Application
Dim AccessDatabase As Database
Dim AccessTable As TableDef
Dim AccessQuery As QueryDef
Dim AccessRecord As Recordset
Dim AccessRecordClone As Recordset
Dim QueryDefName As String

'Grab the access application.
Set AccessApp = Application

'Grab the Current Database in our application.
Set AccessDatabase = AccessApp.CurrentDb

QueryDefName = "PullStockPricesTest"

'Call the `IsInCollection` to see if the QueryDefinition Object Exists.
If IsInCollection(ObjectName:=QueryDefName, CollectionToCheck:=AccessDatabase.QueryDefs) = False Then
    
    'If it doesn't create a new query definition.
    Set AccessQuery = AccessDatabase.CreateQueryDef(Name:=QueryDefName, SQLText:="SELECT * FROM StockPrices")
    
    'Refresh the QueryDefs Collection.
    AccessDatabase.QueryDefs.Refresh
    
    'Refresh the main Database Window.
    AccessApp.RefreshDatabaseWindow
    
Else
    
    'If it does exist, then grab it from the query definitions collection.
    Set AccessQuery = AccessDatabase.QueryDefs(QueryDefName)
    
End If

'Grab when it was last updated.
Debug.Print AccessQuery.LastUpdated

'Grab the SQL Command.
Debug.Print AccessQuery.SQL

'We can even update the SQL Statement after creating a Def.
AccessQuery.SQL = "SELECT * FROM StockPrices ORDER BY date DESC"

'Is it Updatable?
Debug.Print AccessQuery.Updatable

'Once you've defined the query you can open a new recordset.
Set AccessRecord = AccessQuery.OpenRecordset

'Print out the number of records.
Debug.Print "There are " + CStr(AccessRecord.RecordCount) + " records."

'Start working with the recordset object.
With AccessRecord

    'As long as their is a record keep going.
    Do Until .EOF

        'Print out the first column.
        Debug.Print ![Date]
        Debug.Print ![Close]

        'VERY IMPORTANT!!!! IF YOU MISS THIS YOU WILL HAVE A NEVER ENDING LOOP
        .MoveNext

    Loop

End With

'Clone the Record.
Set AccessRecordClone = AccessRecord.Clone
    Debug.Print "Cloned Record Count is: " + CStr(AccessRecordClone.RecordCount)

'Close the Recordset object.
AccessRecord.Close

End Sub


Function IsInCollection(ObjectName As String, CollectionToCheck As Object)

'Declare Variable.
Dim CollectionObject As Object
Dim WasFound As Boolean

'Loop through the "Collection".
For Each CollectionObject In CollectionToCheck
    
    'If the Name matches, it means we have a match.
    If CollectionObject.Name = ObjectName Then
        WasFound = True
    End If

Next

'If it was found we are good.
If WasFound <> True Then
    WasFound = False
End If

'Otherwise we don't have a match.
IsInCollection = WasFound

End Function
