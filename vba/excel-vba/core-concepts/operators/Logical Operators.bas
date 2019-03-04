Attribute VB_Name = "LogicalOperators"
Sub Logical_Operators()
   
   Dim a As Integer
   Dim b As Integer
   
   a = 100
   b = 0
     
   'The AND operator
   If a <> 0 And b <> 0 Then
      MsgBox ("AND Operator Result is : True")
   Else
      MsgBox ("AND Operator Result is : False")
   End If

   'The OR operator
   If a <> 0 Or b <> 0 Then
      MsgBox ("OR Operator Result is : True")
   Else
      MsgBox ("OR Operator Result is : False")
   End If

   'The NOT operator
   If Not (a <> 0 Or b <> 0) Then
      MsgBox ("NOT Operator Result is : True")
   Else
      MsgBox ("NOT Operator Result is : False")
   End If

   'The XOR operator
   If (a <> 0 Xor b <> 0) Then
      MsgBox ("XOR Operator Result is : True")
   Else
      MsgBox ("XOR Operator Result is : False")
   End If
   
   'The EQV operator
   If (a <> 0 Eqv a = 100) Then
      MsgBox ("EQV Operator Result is : True")
   Else
      MsgBox ("EQV Operator Result is : False")
   End If
   
   'The IMP operator
   'If A is true, then B must be true
   If (a <> 0 Imp b = 0) Then
      MsgBox ("IMP Operator Result is : True")
   Else
      MsgBox ("IMP Operator Result is : False")
   End If
     
End Sub
