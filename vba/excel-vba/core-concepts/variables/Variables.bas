Attribute VB_Name = "Variables"
Option Explicit

Public PublicVariable As Integer
Private PrivateVariable As Integer

'Populate the public variable.
Sub PrintPublicVariable()

    PublicVariable = 1000

End Sub

'Populate the private variable.
Sub PrintPrivateVariable()

    PrivateVariable = 2000

End Sub

Sub CallPrivateVariable()
    
    Call PrintPrivateVariable
    MsgBox PrivateVariable
   
End Sub
