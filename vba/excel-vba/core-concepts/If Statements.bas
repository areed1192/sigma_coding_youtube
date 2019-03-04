Sub IfThen()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable equal the value of cell "C3".
    TestScore = Range("C3").Value
    
    'Begin if statement
    If TestScore = 75 Then
       Range("D3").Value = "C"
    End If

End Sub

Sub IfThenElse()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable equal the value of cell "C3".
    TestScore = Range("C3").Value
    
    'Begin if statement
    If TestScore = 75 Then
       Range("D3").Value = "C"
    Else
       Range("D3").Value = "No Score Provided"
    End If

End Sub

Sub IfThenElseIfElse()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable equal the value of cell "C3".
    TestScore = Range("C3").Value
    
    'Begin if statement
    If TestScore = 75 Then
       Range("D3").Value = "C"
    
    ElseIf TestScore = 85 Then
       Range("D3").Value = "B"
    
    Else
       Range("D3").Value = "No Score Provided"
    End If

End Sub

Sub NestedIfStatements()

    'Declare a variable that will house the test score & Favorite Color.
    Dim TestScore As Integer
    Dim FavrColor As String

    'Store values in variables
    TestScore = Range("H3").Value
    FavrColor = Range("G3").Value
    
    'Begin if statement
    If TestScore = 75 Then
        
       'Nested If Statement
       If FavrColor = "Blue" Then
          Range("I3").Value = "Great"
       Else
          Range("I3").Value = "Wonderful"
       End If
    
    ElseIf TestScore = 85 Then
    
       'Nested If Statement
       If FavrColor = "Red" Then
          Range("I3").Value = "Not Great"
       Else
          Range("I3").Value = "Not Wonderful"
       End If
    
    Else
       Range("I3").Value = "No Score Provided"
    End If

End Sub

Sub IfStatementsWithLogicalOperators()

    'Declare a variable that will house the test score & Favorite Color.
    Dim TestScore As Integer
    Dim FavrColor As String
    
    'Store values in variables
    TestScore = Range("H3").Value
    FavrColor = Range("G3").Value
    
    'Begin if statement
    If TestScore = 75 And FavrColor = "Blue" Then
        Range("I3").Value = "Great"
    
    'ElseIf Section One
    ElseIf TestScore = 75 And FavrColor = "Red" Then
        Range("I3").Value = "Wonderful"
    
    'ElseIf Section Two
    ElseIf TestScore = 85 And FavrColor = "Blue" Then
        Range("I3").Value = "Not Great"
    
    'ElseIf Section Three
    ElseIf TestScore = 85 And FavrColor = "Red" Then
        Range("I3").Value = "Not Wonderful"
    
    'Else Section
    Else
        Range("I3").Value = "No Score Provided"
        
    'Close If Block
    End If

End Sub

Sub WrongIfStatement()

    'Declare a variable that will house the test score & Favorite Color.
    Dim TestScore As Integer

    'Store values in variables
    TestScore = Range("H3").Value
    
    'This is incorrect because the Else If will never be reached.
    If TestScore >= 75 Then
       Range("I3").Value = "Great"
    
    ElseIf TestScore >= 85 Then
       'This Portion of the code would never be reached.
       Range("I3").Value = "Wonderful"
        
    End If
    
End Sub

Sub UsingProperties()

    'Declare a variable that will house the test score & Favorite Color.
    Dim Rng As Range

    'Store values in variables
    Set Rng = Range("H3:H5")
    
    'This is incorrect because the Else If will never be reached.
    If Rng.Count = 1 Then
       MsgBox "You have one cell!"
    
    ElseIf Rng.Count > 1 Then
       'This Portion of the code would never be reached.
       MsgBox "You have more than one cell!"
       
    Else
       'This Portion of the code would never be reached.
       MsgBox "Make sure to select a range in order to use this macro."
        
    End If
    
End Sub
