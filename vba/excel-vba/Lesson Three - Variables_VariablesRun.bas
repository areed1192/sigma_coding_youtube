Attribute VB_Name = "VariablesRun"
Option Explicit

'Call my public variable.
Sub CallPublicVariable()
    
    Call PrintPublicVariable
    MsgBox PublicVariable
   
End Sub

'Call my private variable.
Sub CallPrivateVariable()
    
    Call PrintPrivateVariable
    MsgBox PrivateVariable
   
End Sub
