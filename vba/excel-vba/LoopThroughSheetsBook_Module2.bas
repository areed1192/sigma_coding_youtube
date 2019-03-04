Attribute VB_Name = "Module2"
Sub LoopThroughWorksheets_Method1()

    Dim WrkSht As Worksheet
    
    For Each WrkSht In ActiveWorkbook.Worksheets
    
        WrkSht.Range("C3").Value = WrkSht.Name
    
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method2()

    Dim WrkSht As Worksheet
    Dim WrkShtCol As Sheets
    
    Set WrkShtCol = ActiveWorkbook.Worksheets
    
    For Each WrkSht In WrkShtCol
    
        'WrkSht.Range("C3").Value = WrkSht.Name
        WrkSht.Activate
        ActiveCell.Select
        'Application.GoTo WrkSht.ActiveCell.Select
    
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method3()

    Dim WrkSht As Worksheet
    
    For Each WrkSht In ThisWorkbook.Worksheets
    
        WrkSht.Range("C3").Value = WrkSht.Name
    
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method4()

    Dim i As Integer
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
    
        Worksheets(i).Range("C3").Value = "Hello"
    
    Next i
    
End Sub

Sub LoopThroughWorksheets_Method5()

    Dim i As Integer
    
    For i = 1 To ActiveWorkbook.Worksheets.Count Step 2
    
        Worksheets(i).Range("C3").Value = "Hello"
    
    Next i
    
End Sub


Sub LoopThroughWorksheets_Method6()

    Dim i As Integer
    
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
    
        Worksheets(i).Range("C3").Value = "Hello"
    
    Next i
    
End Sub












Sub SelectActiveCellEachWrkSht()

    'Declare your Variables
    Dim WrkSht As Worksheet
    Dim WrkShtCol As Sheets
    
    'Create a Reference to the Worksheets Collection
    Set WrkShtCol = ActiveWorkbook.Worksheets
    
    For Each WrkSht In WrkShtCol
        'Activate the Current Sheet
        WrkSht.Activate
        
        'Select the ActiveCell in that Sheet.
        ActiveCell.Select
    Next WrkSht
    
End Sub








