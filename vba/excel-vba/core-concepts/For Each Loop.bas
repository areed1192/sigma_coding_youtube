'For Each element In Group
'   [statement 1]
'   [statement 2]
'   ....
'   [statement n]
'   [Exit For]
'   [statement 11]
'   [statement 22]
'Next

Sub ForEachLoop()
    
    Dim WrkSht As Worksheet
    
    'Loop through each worksheet (element) in the worksheet collection (group)
    For Each WrkSht In ActiveWorkbook.Worksheets
        Debug.Print WrkSht.Name
    Next WrkSht
    
End Sub

Sub ForEachLoopExit()
    
    Dim WrkSht As Worksheet
    
    'Loop through each worksheet (element) in the worksheet collection (group)
    For Each WrkSht In ActiveWorkbook.Worksheets
        
        Debug.Print WrkSht.Name
        
        'If you found the sheet then exit the for loop
        If WrkSht.Name = "MySheet" Then
           Exit For
        End If
        
    Next WrkSht
    
End Sub

Sub ForEachLoopSimple()
    
    Dim WrkSht As Worksheet
    
    'Loop through each worksheet (element) in the worksheet collection (group)
    For Each WrkSht In ActiveWorkbook.Worksheets
        
        Debug.Print WrkSht.Name
        
        'If you found the sheet then exit the for loop
        If WrkSht.Name = "MySheet" Then
           Exit For
        End If
        
    Next
    
End Sub

Sub NestedForEachLoop()
    
    Dim WrkSht As Worksheet
    Dim Rng As Range
    Dim Cel As Range
    
    'Loop through each worksheet (element) in th worksheet collection (group)
    For Each WrkSht In ThisWorkbook.Worksheets

        'Set a reference to a range.
        Set Rng = WrkSht.Range("A1:B2")
    
            'Loop through the cells in that range.
            For Each Cel In Rng
            
                Cel.Value = 100
                
            Next Cel
    Next WrkSht
    
End Sub
