Sub ComparisonOperators()

'Comparison operators compare two expressions and return a Boolean value that represents the relationship of their values. _
'There are operators for comparing numeric values, operators for comparing strings, and operators for comparing objects.

'Equality Operator
 MsgBox 23 = 33 ' Returns False
 MsgBox 10 = 10 ' Returns True
 
'Inequality Operator
 MsgBox 23 <> 33 ' Returns True
 MsgBox 10 <> 10 ' Returns False
 
'Less Than Operator
 MsgBox 23 < 33 ' Returns True
 MsgBox 10 < 10 ' Returns False
 MsgBox 23 < 12 ' Returns False
 
'Greater Than Operator
 MsgBox 23 > 33 ' Returns False
 MsgBox 10 > 10 ' Returns False
 MsgBox 23 > 12 ' Returns True
 
'Less Than Or Equal To Operator
 MsgBox 23 <= 33 ' Returns True
 MsgBox 10 <= 10 ' Returns True
 MsgBox 23 <= 12 ' Returns False

'Greater Than Or Equal To Operator
 MsgBox 23 >= 33 ' Returns False
 MsgBox 10 >= 10 ' Returns True
 MsgBox 23 >= 12 ' Returns True

End Sub

Sub TypeProperty()
     
    Set Rng = Range("A1")
    
    If TypeOf Rng Is Range Then
        MsgBox "You Have a Range"
    End If
    
End Sub
