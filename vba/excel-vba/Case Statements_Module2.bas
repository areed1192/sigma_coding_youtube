Attribute VB_Name = "Module2"

Sub SelectCase()

    'Declaring our variable
    Dim DayOfWeek As String
    
    'Select starting cell of loop
    ActiveSheet.Range("B3").Select
    
    'Keep looping until the cell value is blank
    Do While ActiveCell.Value <> ""
    
        'Assign my variable to activecell value Im on.
        DayOfWeek = ActiveCell.Value
        
        'Create my case Statement
        Select Case DayOfWeek
        
            Case "Monday", "Tuesday", "Wednesday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Friday", "Saturday"
            ActiveCell.Offset(0, 1).Value = "Neutral"
            
            Case "Sunday", "Thursday"
            ActiveCell.Offset(0, 1).Value = "Bad"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "No Data Provided"
            
        
        End Select
        
        'Go to the next cell
        ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub



Sub SelectCaseRange()

    'Declaring our variable
    Dim Score As String
    
    'Select starting cell of loop
    ActiveSheet.Range("D3").Select
    
    'Keep looping until the cell value is blank
    Do While ActiveCell.Value <> ""
    
        'Assign my variable to activecell value Im on.
        Score = ActiveCell.Value
        
        'Create my case Statement
        Select Case Score
        
            Case Is > 90
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case 79 To 90
            ActiveCell.Offset(0, 1).Value = "Neutral"
            
            Case 69 To 78
            ActiveCell.Offset(0, 1).Value = "Bad"
            
            Case Is < 68
            ActiveCell.Offset(0, 1).Value = "Really Bad"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "No Data Provided"
            
        
        End Select
        
        'Go to the next cell
        ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub


Sub SelectCaseRangeString()

    'Declaring our variable
    Dim DayOfWeek As String
    
    'Select starting cell of loop
    ActiveSheet.Range("B3").Select
    
    'Keep looping until the cell value is blank
    Do While ActiveCell.Value <> ""
    
        'Assign my variable to activecell value Im on.
        DayOfWeek = ActiveCell.Value
        
        'Create my case Statement
        Select Case DayOfWeek
        
            Case "Friday" To "Thursday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "No Data Provided"
            
        
        End Select
        
        'Go to the next cell
        ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub



