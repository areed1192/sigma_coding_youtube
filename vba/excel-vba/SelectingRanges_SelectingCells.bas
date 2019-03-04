Attribute VB_Name = "SelectingCells"
Sub SelectCellOnActiveSheet()

    ActiveSheet.Cells(2, 2).Select
    ActiveSheet.Range("B3").Select
    
End Sub

Sub SelectCellOnDifferentSheet()

    Application.Goto ActiveWorkbook.Sheets("Sheet2").Cells(6, 5)
    Application.Goto (ActiveWorkbook.Sheets("Sheet2").Range("E6"))
    
    'The Long handed way
    'Sheets("Sheet2").Activate
    'ActiveSheet.Cells(6, 5).Select
    
End Sub

Sub SelectCellInDifferentWorkbook()

    Application.Goto Workbooks("BOOK2.xlsx").Sheets("Sheet1").Cells(7, 6)
    Application.Goto Workbooks("BOOK2.xlsx").Sheets("Sheet1").Range("F7")
    
    'The Long handed way
    'Workbooks("BOOK2.xlsx").Sheets("Sheet1").Activate
    'ActiveSheet.Cells(7, 6).Select
    
End Sub

Sub SelectRangeOfCells()

    ActiveSheet.Range(Cells(2, 3), Cells(10, 4)).Select
    ActiveSheet.Range("C2:D10").Select
    ActiveSheet.Range("C2", "D10").Select

End Sub

Sub SelectRangeOfCellsDifferentSheet()

    Application.Goto ActiveWorkbook.Sheets("Sheet3").Range("D3:E11")
    Application.Goto ActiveWorkbook.Sheets("Sheet3").Range("D3", "E11")
    
    'The Long handed way
    'Sheets("Sheet3").Activate
    'ActiveSheet.Range(Cells(3, 4), Cells(11, 5)).Select

End Sub

Sub SelectRangeOfCellsDifferentWorkbook()

   Application.Goto Workbooks("BOOK2.xlsx").Sheets("Sheet1").Range("E4:F12")
   Application.Goto Workbooks("BOOK2.xlsx").Sheets("Sheet1").Range("E4", "F12")
    
   'The Long handed way
   'Workbooks("BOOK2.xlsx").Sheets("Sheet1").Activate
   'ActiveSheet.Range(Cells(4, 5), Cells(12, 6)).Select

End Sub

Sub SelectingNamedRange()

    'ActiveSheet
     Range("Test").Select
    'Application.Goto "Test"
    
    'Different Worksheet, Same Workbook
     Application.Goto Sheets("Sheet1").Range("Test2")
    'Sheets("Sheet1").Activate
    'Range("Test2").Select
    
    'Different Worksheet, Different Workbook
    Application.Goto Workbooks("BOOK2.xlsx").Sheets("Sheet2").Range("Test")
    'Workbooks("BOOK2.xlsx").Sheets("Sheet2").Activate
    'ActiveSheet.Range("Test").Select
 
End Sub
