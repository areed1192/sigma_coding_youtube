Sub ShowUserForm()

' THIS IS THE MODULE CODE.

'Define the UserForm Objects
Dim ExportUserForm As ExcelExportForm
Dim DropDownObjects As MSForms.ComboBox
Dim DropDownApplication As MSForms.ComboBox
Dim UserFormControl As MSForms.Control

'Define Objects related to Excel Objects
Dim xlChart As ChartObject
Dim xlTable As ListObject

'Create a new instance of the UserForm
Set ExportUserForm = New ExcelExportForm

'Grab the Object ComboBox
Set DropDownObjects = ExportUserForm.Controls.Item("ObjectCombo")

'Add Chart names to drop down
For Each xlChart In ThisWorkbook.Worksheets("Sheet1").ChartObjects
    DropDownObjects.AddItem pvargItem:=xlChart.Name + "@Sheet1"
Next
 
'Add Table names to drop down.
For Each xlTable In ThisWorkbook.Worksheets("Sheet1").ListObjects
    DropDownObjects.AddItem pvargItem:=xlTable.Name + "@Sheet1"
Next

'Grab the Object ComboBox
Set DropDownApplication = ExportUserForm.Controls.Item("ApplicationCombo")

'Add the Application Names to the List.
DropDownApplication.AddItem pvargItem:="PowerPoint"
DropDownApplication.AddItem pvargItem:="Word"

'Show the UserForm
ExportUserForm.Show
    
End Sub

'----------------------------
' THIS IS THE USERFORM CODE.
'----------------------------

Private Sub CancelButton_Click()

'Hide the userform to give time to unload
Me.Hide

'Pause it for a few milliseconds
DoEvents
For r = 1 To 10000

    'This is a place holder.
    
Next r

'Unload the Excel Form
Unload ExcelExportForm

End Sub

Private Sub ExportButton_Click()

'Declare Object Variables related to the form.
Dim DropDownObjects As MSForms.ComboBox
Dim DropDownApplication As MSForms.ComboBox
Dim CheckBoxLink As MSForms.CheckBox

'Define variables related to string parsing.
Dim AppSelection As String
Dim ObjSelection As String
Dim LnkSelection As String
Dim SplitObj() As String

'Define Object variables related to Excel Export.
Dim WrkSht As Worksheet
Dim xlChart As ChartObject
Dim xlTable As ListObject


'Grab the Object ComboBox
Set DropDownObjects = Me.Controls.Item("ObjectCombo")

    'Grab the Current Object Selection
    ObjSelection = DropDownObjects.Value

'Grab the Application ComboBox
Set DropDownApplication = Me.Controls.Item("ApplicationCombo")

    'Grab the Current Application Selection
    AppSelection = DropDownApplication.Value

'Grab the Link Checkbox
Set CheckBoxLink = Me.Controls.Item("LinkedCheckBox")

    'Grab the status of the check box
    LnkSelection = CheckBoxLink.Value
    
'Split the object selection on a space.
SplitObj = Split(ObjSelection, "@")
ObjName = SplitObj(0)
ShtName = SplitObj(1)

'Define the worksheet that contains the object
Set WrkSht = ThisWorkbook.Worksheets(ShtName)

'Handle the Object selection.
Select Case True

    Case (ObjName Like "*Chart*")
        
        'If its a chart then grab the chart from the chart objects collection.
        Set xlChart = WrkSht.ChartObjects(ObjName)
        
            'Copy it.
            xlChart.Chart.ChartArea.Copy
    
    Case (ObjName Like "*Table*")
    
        'If its a table then grab the table from the list objects collection.
        Set xlTable = WrkSht.ListObjects(ObjName)
            
            'Copy it.
            xlTable.Range.Copy

End Select

'Handle the Application selection.
Select Case AppSelection

    Case "Word"
        
        'If they choose Word then open a new instance of the Word application.
        Dim WrdApp As Word.Application
        Dim WrdDoc As Word.Document
        Dim WrdSel As Word.Range
        
        'Create a new instance of Word and make it visible.
        Set WrdApp = New Word.Application
            WrdApp.Visible = True
            
            'Add a document.
            Set WrdDoc = WrdApp.Documents.Add
            
            'Grab the first paragraph.
            Set WrdSel = WrdDoc.Paragraphs(1).Range
            
            'Handle paste type.
            If LnkSelection = True Then
            
                'Paste as linked OLEObject.
                WrdSel.PasteSpecial DataType:=wdPasteOLEObject, Link:=True
            
            Else
                
                'Paste as a regular OLEObject.
                WrdSel.PasteSpecial DataType:=wdPasteOLEObject
                
            End If
            
    
    Case "PowerPoint"
    
        'If they choose PowerPoint then open a new instance of the PowerPoint application.
        Dim PptApp As PowerPoint.Application
        Dim PptPres As PowerPoint.Presentation
        Dim PptSld As PowerPoint.Slide
        
        'Create a new instance of the PowerPoint Application and make it visible.
        Set PptApp = New PowerPoint.Application
            PptApp.Visible = True
            
            'Add a new presentation.
            Set PptPres = PptApp.Presentations.Add
            
            'Add a new slide.
            Set PptSld = PptPres.Slides.Add(1, ppLayoutBlank)
            
            'Handle Paste Type
            If LnkSelection = True Then
            
                'Paste as a linked OLEObject.
                PptSld.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=True
            
            Else
                
                'Paste as a regular OLEObject
                PptSld.Shapes.PasteSpecial DataType:=ppPasteOLEObject
                
            End If
        
    Case Else
            
        'Else let them know they need to select and application for the script to work.
        MsgBox Prompt:="You did not select an Office application to export to. Please select an Office Applcation", Buttons:=vbInformation, Title:="No Application Selected"
       
End Select

End Sub
