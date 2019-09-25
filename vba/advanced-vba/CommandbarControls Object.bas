Sub ExploringControls()

'Define Command Bar Object Variables
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

' Define Command Control Object Variables
Dim CommBarControls As CommandBarControls
Dim CommBarControl As CommandBarControl


' Understanding the CommandBarControls Collection
' ----------------------------------------
' We understand that the command bar is simply a group of controls. When we want to work with CommandBarControls we have to first specify the command bar we want to
' work with. After we've done that we can grab that commandbar's controls collection which contains all the CommandBarControls that exist in that CommandBar. Like
' a regular command bar, controls come in different flavors ranging from buttons to drop downs. Additionally, these controls can have different actions and attributes.
'


'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

'This will grab the Cell command bar, the Cell command bar has an index of 38 in the collection.
Set CommBarItem = CommBarColl.Item(38)


'Lets see all the controls in the Cell Command Bar.
Set CommBarControls = CommBarItem.Controls

For Each CommBarControl In CommBarControls

    'Grab some details about each control in the command bar.
    Debug.Print "This control has a caption of: " + CStr(CommBarControl.Caption)
    Debug.Print "This control has an ID of: " + CStr(CommBarControl.ID)
    Debug.Print "This control has an Index of: " + CStr(CommBarControl.Index)
    Debug.Print "This control has a Visible Status of: " + CStr(CommBarControl.Visible)
    Debug.Print "This control has an Enabled Status of: " + CStr(CommBarControl.Enabled)
    Debug.Print "This control has a Type of: " + CStr(CommBarControl.Type)
    Debug.Print "---------------------------------------------------------"

    'Full List of CommandBarControls Object Types can be found here:
    'https://docs.microsoft.com/en-us/office/vba/api/office.msocontroltype


    'Notice that not every control is visible or enabled. This means we can modify the command bar.

Next

End Sub

Sub ExploringControlsDataDump()

'Define Command Bar Object Variables
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

'Define Command Control Object Variables
Dim CommBarControls As CommandBarControls
Dim CommBarControl As CommandBarControl

'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

'This will grab the Cell command bar, the Cell command bar has an index of 38 in the collection.
Set CommBarItem = CommBarColl.Item(38)

'Lets see all the controls in the Cell Command Bar.
Set CommBarControls = CommBarItem.Controls

'Add the headers
Cells(1, 7).Value = "Caption"
Cells(1, 8).Value = "Type"
Cells(1, 9).Value = "Index"
Cells(1, 10).Value = "Visible Status"
Cells(1, 11).Value = "Enabled Status"
Cells(1, 12).Value = "Priority"

'Initalize the row count, start at 2 since we already have headers.
Count = 2

'Loop through each control in the command bar.
For Each CommBarControl In CommBarControls

    'Dump the data.
    Cells(Count, 7).Value = CStr(CommBarControl.Caption)
    Cells(Count, 8).Value = CStr(CommBarControl.Type)
    Cells(Count, 9).Value = CStr(CommBarControl.Index)
    Cells(Count, 10).Value = CStr(CommBarControl.Visible)
    Cells(Count, 11).Value = CStr(CommBarControl.Enabled)
    Cells(Count, 12).Value = CStr(CommBarControl.Priority)
    
    'Increment the Count
    Count = Count + 1
    
Next

End Sub


Sub AddingCustomControls()

'Declare Command Bar Object Variables.
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar

'In this case we are going to create a button, so lets declare a CommandBarButton Object.
Dim CommBarButton As CommandBarButton


'---------------------------------------------------------------------------------------------------------------------------------
'
' This Macro will add a custom button to the Excel Ribbon. It's very important to understand that this won't give you much control
' When it comes to modifying the ribbon. Any controls we add to the ribbon will AUTOMATICALLY be assigned under the "Add-ins" tab.
' There is no way to change this from VBA, you would have to use a different language.
'
'---------------------------------------------------------------------------------------------------------------------------------


'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

'Lets add a new command bar, give it a name, a position, whether you want it replace the active menu bar, and if it's going to be temporary.
Set CommBarItem = CommBarColl.Add(Name:="MySpecialCommandBar", Position:=msoBarTop, MenuBar:=False, Temporary:=True)

'With the new command bar lets add a single button control, define it's type, the ID (default is 1), and if it'll be temporary or not.
Set CommBarButton = CommBarItem.Controls.Add(Type:=msoControlButton, ID:=1, Temporary:=True)

'A note on the ID property, if you set it to an existing control it appears to inherit all the details of that control.

'Lets modify the new command bar button
With CommBarButton

    'Define its style, in this case it's just a button with a caption.
    .Style = msoButtonCaption
    
    'Define its caption.
    .Caption = "Run Highlight Macro"
    
    'When the button is clicked it will run the "HightlightCell" macro
    .OnAction = "HighlightCell"
    
    'Add a tooltip
    .TooltipText = "This button, when clicked, will run the HighlightCell Macro."

End With

'Make sure the CommBarItem is visible
CommBarItem.Visible = True

'Additionally, you could add protections if you wanted to.
'CommBarItem.Protection = msoBarNoCustomize

End Sub


Sub DeletingCustomControls()

'Declare Command Bar Object Variables.
Dim CommBarColl As CommandBars
Dim CommBarItem As CommandBar


'Lets grab the command bar collection, this lives under the application object.
Set CommBarColl = Application.CommandBars

'Lets grab the command bar we just added.
Set CommBarItem = CommBarColl.Item("MySpecialCommandBar")

'Delete the entire command bar.
CommBarItem.Delete

'Delete a specific control from the command bar.
'CommBarItem.Controls.Item(1).Delete

End Sub


Sub HighlightCell()

    'This is the macro I want to run when I click me new button.
    ThisWorkbook.Worksheets("Sheet1").Range("A1").Interior.Color = vbRed

End Sub
