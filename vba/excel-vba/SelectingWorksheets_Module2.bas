Attribute VB_Name = "Module2"
Sub SelectingWorksheets_Method1()

    'Selecting a worksheet using the key method
    'MsgBox Application.ActiveWorkbook.Worksheets("Sara").Name
    MsgBox Worksheets("Sara").Name

End Sub

Sub SelectingWorksheets_Method2()

    'Selecting a worksheet using the index method
     MsgBox Worksheets(3).Name

End Sub


Sub SelectingWorksheets_Method3()

    'Selecting a worksheet using the key method
     MsgBox Sheets("Sara").Name

End Sub

Sub SelectingWorksheets_Method4()

    'Selecting a worksheet using the index method
     MsgBox Sheets(2).Name

End Sub

Sub SelectingWorksheets_Method5()

    'Selecting a worksheet using the activesheet method
     MsgBox ActiveSheet.Name

End Sub

Sub SelectingWorksheets_Method6()

    'Selecting a worksheet using the Code Name method
     MsgBox Sheet1.Name

End Sub
