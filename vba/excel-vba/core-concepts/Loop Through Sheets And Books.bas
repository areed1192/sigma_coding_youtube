Attribute VB_Name = "Module1"
Sub LoopThroughWorksheets_Method1()

    ' --------------------------------------------------------------------------------------------------------------
        'If I want to declare a collection variable I have to declare it as a "SHEETS" collection object.
        'This is because the "WORKSHEETS" property that is returned from a "WORKBOOK" object is a "SHEETS"
        'collection, not a "WORKSHEETS" collection. A "SHEETS" collection object CONTAINS BOTH WORKSHEET OBJECTS
        '& CHART SHEET OBJECTS.
    ' --------------------------------------------------------------------------------------------------------------

    Dim WrkShtCol As Sheets
    Dim WrkSht As Worksheet
    
    Set WrkShtCol = ActiveWorkbook.Worksheets
    
    For Each WrkSht In WrkShtCol
    
        WrkSht.Range("A1").Value = WrkSht.Name
        
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method2()

    Dim WrkSht As Worksheet

    For Each WrkSht In ActiveWorkbook.Worksheets '<<< THIS IS STILL RETURNING A SHEETS COLLECTION OBJECT
    
        WrkSht.Range("A1").Value = WrkSht.Name
        
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method3()

  ' --------------------------------------------------------------------------------------------------------------
    'The "THISWORKBOOK" method will only work in the WORKBOOK THAT CONTAINS THE CODE.
    'For example, if I try to run this code while it is in my PERSONAL MACRO WORKBOOK, it will not work.
    'This is because the "THISWORKBOOK" property returns a "WORKBOOK" object that represents the
    'workbook where the CURRENT MACRO CODE IS RUNNING FROM.
  ' --------------------------------------------------------------------------------------------------------------
  
    Dim WrkSht As Worksheet
    
    For Each WrkSht In ThisWorkbook.Worksheets
        
        With WrkSht.Range("A1")
            .Value = WrkSht.Name
            .Font.Bold = True
        End With
        
    Next WrkSht
    
End Sub

Sub LoopThroughWorksheets_Method4()
    
    Dim i As Integer
    
    For i = 1 To ActiveWorkbook.Worksheets.Count
        
        Worksheets(i).Range("A1").Value = "Hello"
        
    Next i
    
End Sub


Sub LoopThroughWorksheets_Method5()
    
    Dim i As Integer
    
    For i = ActiveWorkbook.Worksheets.Count To 1 Step -1
        
        Worksheets(i).Range("A1").Value = "Hello"
        
    Next i
    
End Sub

Sub LoopThroughWorksheets_Method6()
    
    Dim i As Integer
    
    For i = 1 To ActiveWorkbook.Worksheets.Count Step 2
        
        Worksheets(i).Range("A1").Value = "Hello"
        
    Next i
    
End Sub
