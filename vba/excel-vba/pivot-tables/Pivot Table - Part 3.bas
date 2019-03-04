Sub PivotTablePart3()

'Declare Variables
Dim PvtTbl As PivotTable
Dim PvtFlds As PivotFields
Dim PvtFld As PivotField

'Reference the Pivot Table
Set PvtTbl = ActiveSheet.PivotTables("MyNewPivotTable")

'Loop through each field in our collection
For Each PvtFld In PvtTbl.VisibleFields
    Debug.Print PvtFld.Name
    Debug.Print PvtFld.Value
    Debug.Print PvtFld.DataType
Next

'Create a reference to a single pivot field
Set PvtFld = PvtTbl.PivotFields("Year")
    PvtFld.DragToColumn = False
    PvtFld.DragToData = True
    PvtFld.DragToHide = False
   
'Get the detail for the cell
ActiveCell.ShowDetail = True

'Declare more variables
Dim PvtAxs As PivotAxis
Dim PvtLns As PivotLines
Dim PvtLin As PivotLine

'Reference an Axis of the Pivot Table
Set PvtAxs = PvtTbl.PivotRowAxis
Set PvtLns = PvtAxs.PivotLines

'Loop through each line
For Each PvtLin In PvtLns
    Debug.Print PvtLin.Position
    Debug.Print PvtLin.LineType
    Debug.Print PvtLin.Application
Next

'xlPivotLineBlank        3   Blank line after each group.
'xlPivotLineGrandTotal   2   Grand Total line.
'xlPivotLineRegular      0   Regular PivotLine with pivot items.
'xlPivotLineSubtotal     1   Subtotal line.

'Declare Variable
Dim PvtCell As PivotCell

'Reference an individual Line
Set PvtLin = PvtLns.Item(3)

'Loop through each Pivot cell in our line
For Each PvtCell In PvtLin.PivotLineCellsFull
    Debug.Print PvtCell.Range
    Debug.Print PvtCell.PivotTable
    Debug.Print PvtCell.PivotCellType
    Debug.Print PvtCell.PivotItem
    Debug.Print PvtCell.PivotField
Next

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

Debug.Print PvtTbl.PivotValueCell(3, 1).PivotCell.PivotCellType
Debug.Print PvtTbl.PivotValueCell(3, 1).Value

PvtTbl.PivotValueCell(3, 1).ShowDetail

End Sub
