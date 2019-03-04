Attribute VB_Name = "SelectingLast"

Sub SelectingLastCellOfContiguousRange()

    'Go To last cell
    ActiveSheet.Range("a1").End(xlDown).Select

End Sub

Sub SelectingBlankCellOfContiguousRange()

    'Go to first blank cell after laste cell
    ActiveSheet.Range("a1").End(xlDown).Offset(1, 0).Select
    
End Sub

Sub SelectEntireRangeofContiguousCells()

    'Select Range Of Cells No Blanks
    ActiveSheet.Range("a1", ActiveSheet.Range("a1").End(xlDown)).Select
    ActiveSheet.Range("a1:" & ActiveSheet.Range("a1").End(xlDown).Address).Select
    
End Sub

Sub SelectEntireRangeofNonContiguousCells()
    
    'Select Range of Cells That Includes Blanks
    ActiveSheet.Range("a1", ActiveSheet.Range("a65536").End(xlUp)).Select
    ActiveSheet.Range("a1:" & ActiveSheet.Range("a65536").End(xlUp).Address).Select

End Sub

Sub SelectEntire()

    'Select Entire Row
    Range("1:1").Select
    
    'Select Entire Column
    Range("A:A").Select

End Sub


Sub SelectRectangularRange()

'Select Current Region
'ActiveSheet.Range("a1").CurrentRegion.Select
'ActiveSheet.Range("a1", ActiveSheet.Range("a1").End(xlDown).End(xlToRight)).Select
'ActiveSheet.Range("a1:" & ActiveSheet.Range("a1").End(xlDown).End(xlToRight).Address).Select

'Build Current Region
lastCol = ActiveSheet.Range("a1").End(xlToRight).Column
lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
ActiveSheet.Range("a1", ActiveSheet.Cells(lastRow, lastCol)).Select

'Including a blank row
lastCol = ActiveSheet.Range("a1").End(xlToRight).Column
lastRow = ActiveSheet.Cells(65536, lastCol).End(xlUp).Row
ActiveSheet.Range("a1:" & ActiveSheet.Cells(lastRow, lastCol).Address).Select

End Sub


Sub SelectMultiNonContColumns()

StartRange = "A1"
EndRange = "C1"
Set a = Range(StartRange, Range(StartRange).End(xlDown))
Set b = Range(EndRange, Range(EndRange).End(xlDown))
Union(a, b).Select

End Sub
