Attribute VB_Name = "VariablesStatic"
Option Explicit

Sub Static_Variables()

    Dim NumOne As Integer
    Static NumTwo As Integer
    
    NumOne = NumOne + 1
    NumTwo = NumTwo + 1
    
    MsgBox "Number One Value: " & NumOne
    MsgBox "Number Two Value: " & NumTwo

End Sub

'Run to see static versus dim variable.
Sub Run_Static_Procedure()

    Call Static_Variables

End Sub
