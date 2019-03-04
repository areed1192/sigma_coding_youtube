Attribute VB_Name = "Module1"
Sub ManipulateShapeInPowerPoint()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Dim Excel Variables
    Dim Chrt As ChartObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        PPTApp.Activate
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutTitleOnly)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Object").ChartObjects(1)
        
        'Copy the Chart Object variable we specified above.
        Chrt.Copy
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste
    
    'Set a reference to the shape we want to manipulate
    Set PPTShape = PPTSlide.Shapes(2)
    'Set PPTShape = PPTSlide.Shapes(PPTSlide.Shapes.Count)
   
    'Set the height, width, top & left of the shape
    With PPTShape
        .Left = 100
        .Top = 100
        .Height = 300
        .Width = 300
    End With

End Sub


Sub AligningShapesInPowerPointAlign_MethodOne()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Dim Excel Variables
    Dim Chrt As ChartObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        PPTApp.Activate
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutTitleOnly)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Object").ChartObjects(1)
        
        'Copy the Chart Object variable we specified above.
        Chrt.Copy
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste
    
    'Set a reference to the shape we want to manipulate
    Set PPTShape = PPTSlide.Shapes(PPTSlide.Shapes.Count)
   
    'Set the height & width of the shape
    With PPTShape
        .Height = 300
        .Width = 300
    End With
    
    'Get the Slide Width & Slide Height
    SldHeight = PPTPres.PageSetup.SlideHeight
    SldWidth = PPTPres.PageSetup.SlideWidth
    
    'Calculate the Slide Center and Slide Middle
    SldMiddle = (SldHeight / 2)
    SldCenter = (SldWidth / 2)
    
    'Calculate Shape Center and Shape Middle
    ShpMiddle = (PPTShape.Height / 2)
    ShpCenter = (PPTShape.Width / 2)
    
    'Center the shape and align it to the middle
    With PPTShape
        .Left = SldCenter - ShpCenter
        .Top = SldMiddle - ShpMiddle
    End With
    

End Sub



Sub AligningShapesInPowerPointAlign_MethodTwo()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim PPTShapeRng As PowerPoint.ShapeRange
    Dim ShpCount As Integer
    
    'Dim Excel Variables
    Dim Chrt As ChartObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        PPTApp.Activate
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutTitleOnly)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Object").ChartObjects(1)
        
        'Copy the Chart Object variable we specified above.
        Chrt.Copy
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste
    
        'Count Shapes on Slide
        ShpCount = PPTSlide.Shapes.Count
    
    'Create a reference to a shape range that will contain multiple shapes.
    Set PPTShapeRng = PPTSlide.Shapes.Range(Array(ShpCount))
    
    'Set the height & width of the shape.
    With PPTShapeRng
        .Height = 300
        .Width = 300
    End With
    
    'Align Shape to the middle & center of the SLIDE
    PPTShapeRng.Align msoAlignCenters, True
    PPTShapeRng.Align msoAlignMiddles, True

End Sub


Sub AligningShapesInPowerPointAlign_MethodThree()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide

    'Dim Excel Variables
    Dim Chrt As ChartObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        PPTApp.Activate
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutTitleOnly)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Object").ChartObjects(1)
        
        'Copy the Chart Object variable we specified above.
        Chrt.Copy
   
    'Paste the Chart Object on the Slide that we created above & select that shape
    PPTSlide.Shapes.Paste.Select
    
    'Set dimensions of the shape
    With PPTApp.ActiveWindow.Selection.ShapeRange

        .Height = 300
        .Width = 300
        .Align msoAlignMiddles, True
        .Align msoAlignCenters, True
        
    End With


End Sub





    
