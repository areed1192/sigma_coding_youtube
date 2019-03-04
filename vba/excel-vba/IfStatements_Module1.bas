Attribute VB_Name = "Module1"
Option Explicit

Sub IfThen()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable
    TestScore = 75
    
    'Begin if statement
    If TestScore = 75 Then
       MsgBox "You Got a C"
    End If

End Sub


Sub IfThenElseIf()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable
    TestScore = 85
    
    'Begin if statement
    If TestScore = 75 Then
       MsgBox "You Got a C"
       
    ElseIf TestScore = 85 Then
       MsgBox "You Got a B"
       
    End If

End Sub

Sub IfElse()

    'Declare a variable that will house the test score.
    Dim TestScore As Integer

    'Set the variable
    TestScore = 85
    
    'Begin if statement
    If TestScore = 75 Then
       MsgBox "You Got a C"
       
    ElseIf TestScore = 85 Then
       MsgBox "You Got a B"
    
    Else
       MsgBox "You didn't provide a score."
       
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
    TestScore = 75
    FavrColor = "Blue"
    
    'Begin if statement
    If TestScore = 75 And FavrColor = "Blue" Then
        MsgBox "You have a test score of 75 and your favorite color is blue."
    
    'ElseIf Section Two
    ElseIf TestScore = 85 And FavrColor = "Blue" Then
        MsgBox "You have a test score of 85 and your favorite color is blue."
    
    'ElseIf Section Three
    ElseIf TestScore = 85 And FavrColor = "Red" Then
        MsgBox "You have a test score of 85 and your favorite color is red."
    
    'Else Section
    Else
        MsgBox "You didn't provide the right information."
        
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

