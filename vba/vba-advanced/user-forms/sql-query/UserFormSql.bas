Attribute VB_Name = "UserFormSql"
Sub UserFormSQLShow()

' This subroutine will create a new instance of our UserForm, and format the ListView control so that
' it can be used in an interactive manner. A side note, to use Intellisense with the ListView control
' object we need to enable it's reference doing the following steps:
'
'   Step 1: Select our ToolBox Control Dispaly.
'   Step 2: Right click the display, and left-click "Additional Controls"
'   Step 3: Select "Microsoft ListView Control, version 6.0"
'   Step 4: By selecting that control, it should enable the "Microsoft Windows Common Controls 6.0 (SP6)"
'           reference by default. If it does not, you'll need to go into the Reference libraries and
'           enable it manually.
'
' Unfortunately, I'm not entirely sure if there is official VBA documentation for this control provided by
' Microsoft. However, there is some documentation to the C# version of this control and it can be found at
' this link:
'
'       Link:
'       https://docs.microsoft.com/en-us/dotnet/framework/winforms/controls/listview-control-windows-forms

'Define object variables related to User Form
Dim SqlForm As SQLQueryForm
Dim UserFormControl As MSForms.Control
Dim ListViewControl As ListView

'Create a new instance of the Form.
Set SqlForm = New SQLQueryForm
    
'Grab the QueryResults control, this is the list view object.
Set ListViewControl = SqlForm.Controls("QueryResults")

    'Lets change how our list view behaves.
    
    'I want to be able to reorder columns.
    ListViewControl.AllowColumnReorder = True
    
    'I want the appearance to be flat.
    ListViewControl.Appearance = ccFlat
    
    'I want to be able to select multiple items.
    ListViewControl.MultiSelect = True
    
    'I want a certain font.
    ListViewControl.Font.Name = "Roboto"
    
    'If I select a field it'll select the entire row.
    ListViewControl.FullRowSelect = True
    
    'I want it in a report fashion.
    ListViewControl.View = lvwReport
    
    'I want gridlines.
    ListViewControl.Gridlines = True
    
'Make sure I can select cells behind the form. In other words I want to click outside the User Form.
SqlForm.Show 0

End Sub

