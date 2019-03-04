Sub CreatePivotTable()

'Declare Variables
Dim PvtCache As PivotCache
Dim PvtTbl As PivotTable
Dim PvtFld As PivotField
Dim DataTbl As ListObject

'Delete Pivot Table
ActiveSheet.PivotTables("MyNewPivotTable").TableRange2.Delete

'Create a reference to the data source
Set DataTbl = Worksheets("Data_Table").ListObjects(1)

'Create the Pivot Cache
Set PvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                               SourceData:=DataTbl.Name, _
                                               Version:=6)

'Create the Pivot Table
Set PvtTbl = PvtCache.CreatePivotTable(TableDestination:="Pivot_Table!R1C1", _
                                       TableName:="MyNewPivotTable", _
                                       DefaultVersion:=6)

'Create a Row Field
With PvtTbl.PivotFields("Year")
    .Orientation = xlRowField
    .Position = 1
End With

'Create Another Row Field, this will be the inner one.
With PvtTbl.PivotFields("Country")
    .Orientation = xlRowField
    .Position = 2
End With

'Create a Column Field
With PvtTbl.PivotFields("Month Name")
    .Orientation = xlColumnField
    .Position = 1
End With

'Create a Data Field
With PvtTbl.PivotFields("COGS")
    .Orientation = xlDataField
    .Position = 1
End With

'Create a Filter Field
With PvtTbl.PivotFields("Product")
    .Orientation = xlPageField
    .Position = 1
End With

'Hide a field
Set PvtFld = PvtTbl.PivotFields("Month Name")
    PvtFld.PivotItems("January").Visible = False

'Change the Layout & Style
PvtTbl.RowAxisLayout xlTabularRow
PvtTbl.TableStyle2 = "PivotStyleLight24"

End Sub
