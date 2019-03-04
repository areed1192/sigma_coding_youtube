Attribute VB_Name = "Practice"
Option Explicit

Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
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



'PvtTbl.PivotValueCell(10, 1).PivotCell
'PvtTbl.PivotValueCell(10, 1).Value
'PvtTbl.PivotValueCell(10, 1).ShowDetail

Sub Part3()

Dim PvtTbl As PivotTable
Dim PvtFields As PivotFields
Dim PvtField As PivotField

'Reference the Pivot Table
Set PvtTbl = ActiveSheet.PivotTables("MyNewPivotTable")

'Referencing Fields
'PvtTbl.DataFields
'PvtTbl.RowFields
'PvtTbl.ColumnFields
'PvtTbl.PageFields
'PvtTbl.CubeFields
'PvtTbl.HiddenFields
'PvtTbl.VisibleFields

'For Each PvtField In PvtTbl.VisibleFields
'    Debug.Print PvtField.Name
'    Debug.Print PvtField.Value
'Next PvtField

'Set PvtField = PvtTbl.PivotFields("Year")
    'PvtField.AutoGroup
    'PvtField.DragToColumn = False
    'PvtField.DragToData = True
    'PvtField.DragToHide = False


'Debug.Print Range("C10").PivotCell.PivotCellType

'PvtTbl.PivotSelect "Year['2014'] Country['Canada'] Month Name['February'] Values['Sum of COGS']", xlDataOnly
'ActiveCell.ShowDetail = True


Dim PvtAxs As PivotAxis
Dim PvtLns As PivotLines
Dim PvtLin As PivotLine

Set PvtAxs = PvtTbl.PivotRowAxis
Set PvtLns = PvtAxs.PivotLines

'For Each PvtLin In PvtLns
'    Debug.Print PvtLin.Position
'    Debug.Print PvtLin.LineType
'Next

'xlPivotLineBlank        3   Blank line after each group.
'xlPivotLineGrandTotal   2   Grand Total line.
'xlPivotLineRegular      0   Regular PivotLine with pivot items.
'xlPivotLineSubtotal     1   Subtotal line.

Dim PvtCell As PivotCell

Set PvtLin = PvtLns.Item(3)
   

'For Each PvtCell In PvtLin.PivotLineCellsFull
'    Debug.Print PvtCell.Range
'    Debug.Print PvtCell.PivotTable
'    Debug.Print PvtCell.PivotItem
'    Debug.Print PvtCell.PivotField
'    Debug.Print PvtCell.PivotCellType
'    'Debug.Print PvtCell.ServerActions
'Next

'xlPivotCellBlankCell        9   A structural blank cell in the PivotTable.
'xlPivotCellCustomSubtotal   7   A cell in the row or column area that is a custom subtotal.
'xlPivotCellDataField        4   A data field label (not the Data button).
'xlPivotCellDataPivotField   8   The Data button.
'xlPivotCellGrandTotal       3   A cell in a row or column area that is a grand total.
'xlPivotCellPageFieldItem    6   The cell that shows the selected item of a Page field.
'xlPivotCellPivotField       5   The button for a field (not the Data button).
'xlPivotCellPivotItem        1   A cell in the row or column area that is not a subtotal, grand total, custom subtotal, or blank line.
'xlPivotCellSubtotal         2   A cell in the row or column area that is a subtotal.
'xlPivotCellValue            0   Any cell in the data area (except a blank row).


End Sub


