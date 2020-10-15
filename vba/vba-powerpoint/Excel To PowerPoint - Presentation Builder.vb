Option Explicit

Sub ExportTable()

Dim xlBook As Workbook
Dim xlSheet As Worksheet
Dim xlTable As ListObject
Dim xlTableColumn As ListColumn
Dim xlTableRow As Range

Dim xlChartObject As ChartObject
Dim xlTableObject As ListObject

Dim pptApp As PowerPoint.Application
Dim pptPres As PowerPoint.Presentation
Dim pptSlide As PowerPoint.Slide
Dim pptShape As PowerPoint.Shape

Dim ObjectField As String
Dim ObjectFieldParts As Variant

Dim ObjectName As String
Dim ObjectSheet As String
Dim ObjectType As String

'Set The Book.
Set xlBook = ThisWorkbook

'Set the Sheet.
Set xlSheet = xlBook.Worksheets("Export")

'Grab the Table.
Set xlTable = xlSheet.ListObjects("ExportToPowerPoint")

'Create an Instance of PowerPoint.
On Error Resume Next

    'Grab the Active PowerPoint Application, if it's there.
    Set pptApp = GetObject(, "PowerPoint.Application")

        'If the Application isn't open it will return a 429 error
        If Err.Number = 429 Then

            'If it is not open then clear the error and create a new instance of PowerPoint
            Err.Clear
            Set pptApp = New PowerPoint.Application

                'Make it Visible.
                pptApp.Visible = True

                'Bring it to the front.
                pptApp.Activate

        End If

'Create the Presentation.
Set pptPres = pptApp.Presentations.Add

'Loop through each table.
For Each xlTableRow In xlTable.ListColumns("Object").DataBodyRange

    'Grab the Object Field.
    ObjectField = xlTableRow.Value
    
    'Console Log.
    Debug.Print "Exporting Object: " + ObjectField
    
    'Split it.
    ObjectFieldParts = Split(ObjectField, "-")
    
    'Grab the Object Name.
    ObjectName = ObjectFieldParts(0)
    
    'Grab the Object Sheet Location.
    ObjectSheet = ObjectFieldParts(1)
    
    'Grab the Object Type Name.
    ObjectType = ObjectFieldParts(2)
    
    'Grab the Sheet.
    Set xlSheet = xlBook.Worksheets(ObjectSheet)
        xlSheet.Activate
    
    'Grab the Object.
    If ObjectType = "ListObject" Then
        
        'Copy it.
        Set xlTableObject = xlSheet.ListObjects.Item(ObjectName)
            xlTableObject.Range.Copy
    
    ElseIf ObjectType = "ChartObject" Then
        
        'Copy it.
        Set xlChartObject = xlSheet.ChartObjects(ObjectName)
            xlChartObject.Chart.ChartArea.Copy
    
    End If
    
    'Pause it.
    Application.Wait Now + #12:00:01 AM#
    
    'Console Log.
    Debug.Print "Adding Slide: " + CStr(xlTableRow.Offset(0, 1).Value)
    
    'Add the Slide.
    Set pptSlide = pptPres.Slides.Add(xlTableRow.Offset(0, 1).Value, ppLayoutTitleOnly)
    
    'Paste the Shape.
    pptSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=xlTableRow.Offset(0, 6).Value
    
    'Change the Slide Title.
    pptSlide.Shapes(1).TextFrame.TextRange = xlTableRow.Offset(0, 7)
    
    'Grab the Shape.
    Set pptShape = pptSlide.Shapes(pptSlide.Shapes.Count)
    
    'Resize the Shape.
    pptShape.Top = xlTableRow.Offset(0, 2).Value
    pptShape.Width = xlTableRow.Offset(0, 3).Value
    pptShape.Left = xlTableRow.Offset(0, 4).Value
    pptShape.Height = xlTableRow.Offset(0, 5).Value

Next

End Sub

Sub UpdateDropdownColumn()

Dim xlBook As Workbook
Dim xlSheet As Worksheet
Dim xlTable As ListObject
Dim xlTableColumn As ListColumn
Dim xlChartObject As ChartObject
Dim xlTableObject As ListObject

Dim ObjectArray() As String
Dim ObjectArrayIndex As Integer

'Set The Book.
Set xlBook = ThisWorkbook

'Loop through each Worksheet.
For Each xlSheet In xlBook.Worksheets
    
    'If we have Charts.
    If xlSheet.ChartObjects.Count > 0 Then
        
        'Grab each Chart Name.
        For Each xlChartObject In xlSheet.ChartObjects
            
            'Update the Count.
            ObjectArrayIndex = ObjectArrayIndex + 1
            ReDim Preserve ObjectArray(ObjectArrayIndex)
            
            'Add the Chart Object to the Array.
            ObjectArray(ObjectArrayIndex) = xlChartObject.Name & "-" & xlSheet.Name & "-" & TypeName(xlChartObject)
            
        Next xlChartObject
        
    End If
    
    'If we have Tables.
    If xlSheet.ListObjects.Count > 0 Then
        
        'Grab each Table Name.
        For Each xlTableObject In xlSheet.ListObjects
        
            'Update the Count.
            ObjectArrayIndex = ObjectArrayIndex + 1
            ReDim Preserve ObjectArray(ObjectArrayIndex)
            
            'Add the Chart Object to the Array.
            ObjectArray(ObjectArrayIndex) = xlTableObject.Name & "-" & xlSheet.Name & "-" & TypeName(xlTableObject)
            
        Next xlTableObject
        
    End If
    
Next xlSheet

'Set the Sheet.
Set xlSheet = xlBook.Worksheets("Export")

'Grab the Table.
Set xlTable = xlSheet.ListObjects("ExportToPowerPoint")

'Grab the Object Column.
Set xlTableColumn = xlTable.ListColumns("Object")

'Set the Validation Dropdown.
With xlTableColumn.DataBodyRange.Validation
    
    'Delete the Old One.
    .Delete
    
    'Add the New Data.
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(ObjectArray, ",")
    
    'Make sure it's a dropdown.
    .InCellDropdown = True
    
End With
 
End Sub