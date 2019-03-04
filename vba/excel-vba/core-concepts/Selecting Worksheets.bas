Attribute VB_Name = "Module1"
Sub SelectWorkSheet_Method1()

    'Worksheets Collection - Key Method

    Worksheets("Sheet2").Range("A1").Value = 300
    
End Sub

Sub SelectWorksheet_Method2()

    'Worksheets Collection - Index Method
    
    Worksheets(2).Range("A1").Value = 300
    
End Sub

Sub SelectWorksheet_Method3()

    'Sheets Collection - Key Method

    Sheets("Sheet1").Range("A1").Value = 300
    
End Sub

Sub SelectWorksheet_Method4()

    'Sheets Collection - Index Method
    
    Sheets(1).Range("A1").Value = 300
    
End Sub

Sub SelectWorksheet_Method5()

    'ActiveSheet Method

    ActiveSheet.Range("A1").Value = 300
    
End Sub

Sub SelectWorksheet_Method6()

    'Unqualified Object Name Method (Also Called CodeName Method)
    'WARNING - UNSTABLE WHEN YOU HAVE STORED IN A PERSONAL MACRO WORKBOOK
    
    Sheet1.Range("A1").Value = 300
    
End Sub
