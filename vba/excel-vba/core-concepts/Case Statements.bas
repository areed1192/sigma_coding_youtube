Attribute VB_Name = "Module1"
Option Explicit

Sub CaseStatements()

    'Declare Variable
    Dim DayOfWeek As String
    
    'Select the starting cell
    ActiveSheet.Range("B3").Select
    
    'Create a do while loop that will loop through the range of cells.
    Do While ActiveCell.Value <> ""
       
       'Set my variable equal to active cell value on each iteration of the loop
       DayOfWeek = ActiveCell.Value
       
       'Create the case statement
       Select Case DayOfWeek
              
            Case "monday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Tuesday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Wednesday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Thursday"
            ActiveCell.Offset(0, 1).Value = "Bad"
            
            Case "Friday"
            ActiveCell.Offset(0, 1).Value = "Bad"
            
            Case "Saturday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Sunday"
            ActiveCell.Offset(0, 1).Value = "Neutral"
       
       End Select
       
       'GO TO THE NEXT CELL --- VERY IMPORTANT
       ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub

Sub CaseStatements_Operator()

    'Declare Variable
    Dim DayOfWeek As String
    
    'Select the starting cell
    ActiveSheet.Range("B3").Select
    
    'Create a do while loop that will loop through the range of cells.
    Do While ActiveCell.Value <> ""
       
       'Set my variable equal to active cell value on each iteration of the loop
       DayOfWeek = ActiveCell.Value
       
       'Create the case statement
       Select Case DayOfWeek
              
            Case "Monday", "Tuesday", "Wednesday", "Saturday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Thursday", "Friday"
            ActiveCell.Offset(0, 1).Value = "Bad"
           
            Case "Sunday"
            ActiveCell.Offset(0, 1).Value = "Neutral"
       
       End Select
       
       'GO TO THE NEXT CELL --- VERY IMPORTANT
       ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub


Sub CaseStatements_Else()

    'Declare Variable
    Dim DayOfWeek As String
    
    'Select the starting cell
    ActiveSheet.Range("B3").Select
    
    'Create a do while loop that will loop through the range of cells.
    Do While ActiveCell.Value <> ""
       
       'Set my variable equal to active cell value on each iteration of the loop
       DayOfWeek = ActiveCell.Value
       
       'Create the case statement
       Select Case DayOfWeek
              
            Case "Monday", "Tuesday", "Wednesday", "Saturday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Thursday", "Friday"
            ActiveCell.Offset(0, 1).Value = "Bad"
           
            Case "Sunday"
            ActiveCell.Offset(0, 1).Value = "Neutral"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "You Messed Up Big Time"
       
       End Select
       
       'GO TO THE NEXT CELL --- VERY IMPORTANT
       ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub


Sub CaseStatements_Range()

    'We use IS keyword with comparison operators
    'We use TO keyword with a range of values

    'Declare Variable
    Dim DayOfWeek As String
    Dim DayScore As Variant
    
    'Select the starting cell
    ActiveSheet.Range("D3").Select
    
    'Create a do while loop that will loop through the range of cells.
    Do While ActiveCell.Value <> ""
       
       'Set my variable equal to active cell value on each iteration of the loop
       DayScore = ActiveCell.Value
       
       'Create the case statement
       Select Case DayScore
              
            Case Is > 90
            ActiveCell.Offset(0, 1).Value = "A"
            
            Case 79 To 90
            ActiveCell.Offset(0, 1).Value = "B"
           
            Case 69 To 78
            ActiveCell.Offset(0, 1).Value = "C"
            
            Case 59 To 68
            ActiveCell.Offset(0, 1).Value = "D"
            
            Case Is < 58
            ActiveCell.Offset(0, 1).Value = "F"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "You Messed Up Big Time"
       
       End Select
       
       'GO TO THE NEXT CELL --- VERY IMPORTANT
       ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub


Sub CaseStatements_Strings()

    'Declare Variable
    Dim DayOfWeek As String
    
    'Select the starting cell
    ActiveSheet.Range("B3").Select
    
    'Create a do while loop that will loop through the range of cells.
    Do While ActiveCell.Value <> ""
       
       'Set my variable equal to active cell value on each iteration of the loop
       DayOfWeek = ActiveCell.Value
       
       'Create the case statement
       Select Case DayOfWeek
              
            'Any value that falls within the alphabetical range
            Case "Monday" To "Saturday"
            ActiveCell.Offset(0, 1).Value = "Good"
            
            Case "Sunday"
            ActiveCell.Offset(0, 1).Value = "Ok Good"
            
            Case Else
            ActiveCell.Offset(0, 1).Value = "You Messed Up Big Time"
       
       End Select
       
       'GO TO THE NEXT CELL --- VERY IMPORTANT
       ActiveCell.Offset(1, 0).Select
    
    Loop

End Sub
