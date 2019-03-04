Sub CreatePivotTable()

'Declare Variables
Dim PvtCache As PivotCache
Dim PvtTbl As PivotTable
Dim PvtFld As PivotField
Dim DataTbl As ListObject

'Delete Pivot Table
ActiveSheet.PivotTables("MyNewPivotTable").TableRange2.Delete

'Create a reference to the data table
Set DataTbl = Worksheets("Data_Table").ListObjects(1)

'Create a Pivot Cache
Set PvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                               SourceData:=DataTbl.Name, _
                                               Version:=6)
                                               
'Create a Pivot Table
Set PvtTbl = PvtCache.CreatePivotTable(TableDestination:="Pivot_Table!R1C1", _
                                       TableName:="MyNewPivotTable", _
                                       DefaultVersion:=6)
                                       
'Create a row field
With PvtTbl.PivotFields("Year")
    .Orientation = xlRowField
    .Position = 1
End With

'Create a row field
With PvtTbl.PivotFields("Country")
    .Orientation = xlRowField
    .Position = 2
End With

'Create a column field
With PvtTbl.PivotFields("Month Name")
    .Orientation = xlColumnField
    .Position = 1
End With

'Create a data field
With PvtTbl.PivotFields("COGS")
    .Orientation = xlDataField
    .Position = 1
End With

'Create a filter field
With PvtTbl.PivotFields("Product")
    .Orientation = xlPageField
    .Position = 1
End With

'Hide a pivot item
Set PvtFld = PvtTbl.PivotFields("Month Name")
    PvtFld.PivotItems("March").Visible = False
    
'Change the layout
PvtTbl.RowAxisLayout xlTabularRow

'Change the color
PvtTbl.TableStyle2 = "PivotStyleLight24"

'Create a calculate field
PvtTbl.CalculatedFields.Add "Average Selling Price", "=Gross Sales / Units Sold"

'Create the field to house the calculated field
With PvtTbl.PivotFields("Average Selling Price")
    .Orientation = xlDataField
    .Position = 2
    .NumberFormat = "#,##0.00"
End With

'Add a Calculated Field
PvtTbl.CalculatedFields.Add "Average Selling Price 2", "= Average Selling Price *.1"

'Create a Data Field
With PvtTbl.PivotFields("Average Selling Price 2")
    .Orientation = xlDataField
    .Position = 3
    .NumberFormat = "#,##0.00"
End With

'Add a Calculated Item
PvtTbl.PivotFields("Country").CalculatedItems.Add "North America", "=Canada + Mexico + United States of America"


'Get Pivot Data With VBA
Range("A22").Value = PvtTbl.GetPivotData("Sum Of COGS", "Country", "Canada", "Year", "2014", "Month Name", "April")

'Get Data
Range("A23").Value = PvtTbl.GetData("Sum Of COGS Canada 2014 April")

'Pivot Select VBA
PvtTbl.PivotSelect "Country['France'] Year['2013']", xlDataAndLabel

'Selecting Different Parts of Our Pivot Table
PvtTbl.DataBodyRange.Select
PvtTbl.DataLabelRange.Select
PvtTbl.RowRange.Select
PvtTbl.ColumnRange.Select
PvtTbl.PageRange.Select
PvtTbl.TableRange1.Select
PvtTbl.TableRange2.Select

'Clear all Filters
PvtTbl.PivotFields("Month Name").ClearAllFilters

'Add a Top 2 Filter
PvtTbl.PivotFields("Month Name").PivotFilters.Add2 Type:=xlTopCount, DataField:=PvtTbl.PivotFields("Sum of COGS"), Value1:=2

'Add a Label Filter
PvtTbl.PivotFields("Month Name").PivotFilters.Add2 Type:=xlCaptionContains, Value1:="Feb"

End Sub
