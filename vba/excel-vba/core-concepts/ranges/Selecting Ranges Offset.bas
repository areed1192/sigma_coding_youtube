Attribute VB_Name = "Offset"
Sub OffsetActiveCell()

'Go 5 rows below & 4 columns to the left
ActiveCell.Offset(5, -4).Select

'Go 2 rows above & 3 columns to the right
ActiveCell.Offset(-2, 3).Select

'Error occurs if the row you're selecting is off the sheet.

End Sub

Sub OffsetCell()

'Go 5 rows below & 4 columns to the right
ActiveSheet.Cells(7, 3).Offset(5, 4).Select

'Go 5 rows below & 4 columns to the right
ActiveSheet.Range("C7").Offset(5, 4).Select

End Sub


Sub OffsetRangeOfCell()

'Go 4 rows below & 3 columns to the right - MAINTAING THE SAME RANGE SIZE
ActiveSheet.Range("Test").Offset(4, 3).Select

'Long handed way
'Go 4 rows below & 3 columns to the right - MAINTAING THE SAME RANGE SIZE
Sheets("Sheet2").Activate
ActiveSheet.Range("Test").Offset(4, 3).Select

End Sub

Sub ResizeSelection()

'Select the range
Range("Test").Select

'Resize the selection by five rows
Selection.Resize(Selection.Rows.Count + 5, Selection.Columns.Count).Select

End Sub


Sub ResizeSelectionOffset()

'Select the range
Range("Test").Select

'Offset and then resize the selection by five rows
Selection.Offset(4, 3).Resize(Selection.Rows.Count + 5, Selection.Columns.Count).Select

End Sub


Sub SelectUnionOfTwoOrMoreRanges()

Application.Union(Range("Test"), Range("Sample")).Select

'DOES NOT WORK ACROSS SHEETS
Set y = Application.Union(Range("Sheet1!A1:B2"), Range("Sheet1!C3:D4"))
Set y = Application.Union(Range("Sheet1!A1:B2"), Range("Sheet2!C3:D4"))


End Sub



Sub SelectIntersection()

'DOES NOT WORK ACROSS SHEETS
Application.Intersect(Range("Test"), Range("Sample")).Select

End Sub
