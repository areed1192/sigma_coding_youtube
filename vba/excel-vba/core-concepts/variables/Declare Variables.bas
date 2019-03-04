Attribute VB_Name = "VariablesDeclare"
Option Explicit

Sub Declaring_Variables()

    'The explicit way
    Dim FirstName As String
        FirstName = "Alex"
    
    'The implicit way, but keep in mind it's now variant.
    LastName = "Awesome"
    
    'Using the let statement.
    Let MiddleName = "Cool"

End Sub
