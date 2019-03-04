Attribute VB_Name = "Module3"
Option Explicit

Sub ManipulateShapeInPowerPoint()
        
    'Declare PowerPoint Variables
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide

    'Create a new presentation in the PowerPoint Application
    Set PPTPres = Application.ActivePresentation
    
    'Reference the active slide
    Set PPTSlide = PPTPres.Slides(1)
   
    'Select the shape
    PPTSlide.Shapes.Select
    
    'Set the dimensions of the selected shape in my ACTIVE WINDOW
    With Application.ActiveWindow.Selection.ShapeRange
        .Height = 300
        .Width = 300
        .Align msoAlignCenters, True
        .Align msoAlignMiddles, True
    End With

End Sub
