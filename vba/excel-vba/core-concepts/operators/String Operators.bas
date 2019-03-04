Sub StringOperators()

    'Comments are preceded by the "'" symbol

    Str1 = "Hi"
    Str2 = "There"
        
    'This is how we combine two strings using the "&" and the "+" symbols.
    MsgBox Str1 + " " + Str2
    MsgBox Str1 & " " & Str2
    
    'Here is how write combine multiple lines into a single line using the "&" and "+" symbol.
    MultiLineStr = "Hi There " + _
                   "my name is " + _
                   "Alex"
                   
    MultiLineStr = "Hi There " & _
                   "my name is " & _
                   "Alex"
                   
    MsgBox MultiLineStr 
       
    'Lets see if we can condense this for loop into a single line.
    For i = 1 To 10
    
        x = 1 + 1
    
    Next i
    
    'If we use the ":" symbol we can combine all the different parts into a single line.
    For i = 1 To 10: x = 1 + 1: Next i
    

End Sub
