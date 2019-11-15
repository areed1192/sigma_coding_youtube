'Microsoft Doc CommandBar: https://docs.microsoft.com/en-us/office/vba/api/office.commandbar 
'Microsoft Doc CommandBars: https://docs.microsoft.com/en-us/office/vba/api/office.commandbars 
'Microsoft Doc CommandBarControl: https://docs.microsoft.com/en-us/office/vba/api/office.commandbarcontrol
'Microsoft Doc CommandBarControls: https://docs.microsoft.com/en-us/office/vba/api/office.commandbarcontrols

Sub ExploringCommandBars()
      
'Declare your variables
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

' Understanding the CommandBar Collection
' ----------------------------------------
' The Command bar collection, contains all the command bar objects. A command bar object can be thought of as a collection of controls (buttons, dropdowns, or menus).
' For example, if you're in the Excel worksheet and you right click on a cell then the Cell Command Bar Object will display. There are multiple of these command bar
' objects throughout any of the Office applications.
'
' Generally speaking, there are three type of command bar objects.
'
'       1. A Shortcut Menu (msoBarTypePopup) - 1
'       2. A Default Command Bar (msoBarTypeNormal) - 0
'       3. A Menu Bar (msoBarTypeMenuBar) - 2


'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

    'How many command bars do we have?
    Debug.Print "There are " + CStr(CommBarColl.Count) + " command bar objects in the Excel Application."
    Debug.Print "---------------------------------------------------------------------------------------"
    
    'Lets see their name, their type, their position, and whether they are visible or not. This will require use to loop through each command bar in the collection.
    For Each CommBarItem In CommBarColl
    
        If CommBarItem.Visible = False Then
    
        'Display the info to the user.
        Debug.Print "The Name of this command bar object is: " + CStr(CommBarItem.Name)
        Debug.Print "The Type of command bar object is: " + CStr(CommBarItem.Type)
        Debug.Print "The Position of this command bar object is: " + CStr(CommBarItem.Position)
        Debug.Print "This command bar object has a visible status of: " + CStr(CommBarItem.Visible)
        
        
        'Full List of CommandBar Object Types can be found here:
        'https://docs.microsoft.com/en-us/office/vba/api/office.msobartype
        
        
        'Additionally, you could also determine if the CommandBar Object is built-in or not.
        Debug.Print "This command bar object has a built-in status of: " + CStr(CommBarItem.BuiltIn)
        Debug.Print "---------------------------------------------------------"
        
        End If
        
    Next
    
End Sub


Sub ExploringCommandBarsDataDump()

'Declare your variables
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

    'Add the headers
    Cells(1, 1).Value = "Name"
    Cells(1, 2).Value = "Type"
    Cells(1, 3).Value = "Index"
    Cells(1, 4).Value = "Visible Status"
    Cells(1, 5).Value = "Enabled Status"
    
    'Initalize the row count, start at 2 since we already have headers.
    Count = 2
    
    'Lets dump all the details in the Excel Worksheet
    For Each CommBarItem In CommBarColl

        'Dump the data.
        Cells(Count, 1).Value = CStr(CommBarItem.Name)
        Cells(Count, 2).Value = CStr(CommBarItem.Type)
        Cells(Count, 3).Value = CStr(CommBarItem.Index)
        Cells(Count, 4).Value = CStr(CommBarItem.Visible)
        Cells(Count, 5).Value = CStr(CommBarItem.Enabled)
        
        'Increment the Count
        Count = Count + 1
        
    Next

End Sub


Sub DisplayingACommandBar()

'Declare your variables
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

'This will grab the Cell command bar, the Cell command bar has an index of 38 in the collection.
Set CommBarItem = CommBarColl.Item(38)

'-------------------------------------------------------------------------------------------------
'Note that if the Command Bar Object does not have a type of MsoBarPopUp (2) then the following
'method will fail. I selected the Cell command bar because that has a Type of MsoBarPopUp
'-------------------------------------------------------------------------------------------------

'This will show the pop up, the coordinates are optional.
CommBarItem.ShowPopup x:=200, y:=200

End Sub
