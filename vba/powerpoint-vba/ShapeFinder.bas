'Declare Private Variables
Private ShpLeft, ShpHeight, ShpTop, ShpWidth As Integer
Private PPTShape As Selection

Sub FindDimensions()
    'Declare Our Selection Object, then set it equal to the active selection in our window.
    Set PPTShape = ActiveWindow.Selection
    
    'If the selection is PowerPoint Shape then retrieve the dimensions and store them.
    If PPTShape.Type = ppSelectionShapes Then
        With PPTShape
           ShpHeight = .ShapeRange.Height
           ShpLeft = .ShapeRange.Left
           ShpTop = .ShapeRange.Top
           ShpWidth = .ShapeRange.Width
        End With
    Else
        MsgBox "The object you have selected is not a PowerPoint Shape Object. Please select a shape.", vbCritical, "Macro Error Shape Finder"
    End If
End Sub

Sub ApplyDimensions()
    'Declare Our Selection Object, then set it equal to the active selection in our window.
    Set PPTShape = ActiveWindow.Selection
        
    'If the selection is PowerPoint Shape then apply the dimensions we have stored.
    If PPTShape.Type = ppSelectionShapes Then
        With PPTShape
           .ShapeRange.Height = ShpHeight
           .ShapeRange.Left = ShpLeft
           .ShapeRange.Top = ShpTop
           .ShapeRange.Width = ShpWidth
        End With
    Else
        MsgBox "The object you have selected is not a PowerPoint Shape Object. Please select a shape.", vbCritical, "Macro Error Shape Apply"
    End If
End Sub
