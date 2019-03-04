'BASIC SYNTAX
'For counter = start To end [Step increment]
'   {...statements...}
'Next [counter]


'Parameters or Arguments

'counter
'The loop counter variable.

'Start
'The starting value for counter.

'End
'The ending value for counter.

'increment
'Optional. The value that counter is incremented each pass through the loop.
'It can be a positive or negative number.
'If not specified, it will default to an increment of 1 so that each pass through the loop increases counter by 1.

'statements
'The statements of code to execute each pass through the loop.

Sub ForLoops()

'The basic for loop
For i = 1 To 10
    Cells(i, 1).Value = i
Next i

End Sub

Sub ForLoopsStep()

'The basic for loop
For i = 1 To 10 Step 2
    Cells(i, 2).Value = i
Next i

End Sub

Sub ForLoopsReverse()

'For loop going in reverse order.
For i = 10 To 1 Step -1
    Cells(i, 3).Value = i
Next i

End Sub

Sub NestedLoops()

'Outer Loop
For i = 1 To 3
    'Inner Loop
    For j = 12 To 16
        Cells(j, i).Value = 100
    Next j
Next i

End Sub


Sub ExitForLoop()

For i = 1 To 20

    ' Display the index.
    Debug.Print i
    
    ' If index is 10, exit the loop.
    If i = 10 Then
       Exit For
    End If

Next i

End Sub
