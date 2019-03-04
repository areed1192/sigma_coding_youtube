Attribute VB_Name = "Module2"

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
    
    'Set a reference to the shape I want to manipulate
    Set PPTShape = PPTSlide.Shapes(2)
    
    'Set the left, top, height & width of my shape
    With PPTShape
        .Left = 200
        .Top = 200
        .Height = 300
        .Width = 300
    End With

End Sub

Sub ManipulateShapeInPowerPoint_Method2()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim ShpCnt As Integer
    
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
    
    'Count the number of shapes on my slide
    ShpCnt = PPTSlide.Shapes.Count
    
    'Set a reference to the shape I want to manipulate
    Set PPTShape = PPTSlide.Shapes(ShpCnt)
    
    'Set the left, top, height & width of my shape
    With PPTShape
        .Left = 200
        .Top = 200
        .Height = 300
        .Width = 300
    End With

End Sub

Sub ManipulateShapeInPowerPoint_Method3()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim ShpCnt, SldHeight, SldWidth, ShpMiddle, ShpCenter, SldMiddle, SldCenter As Integer
    
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
    
    'Count the number of shapes on my slide
    ShpCnt = PPTSlide.Shapes.Count
    
    'Set a reference to the shape I want to manipulate
    Set PPTShape = PPTSlide.Shapes(ShpCnt)
    
    'Set the left, top, height & width of my shape
    With PPTShape
        .Height = 300
        .Width = 300
    End With
    
    'Get slide height & slide width
    SldHeight = PPTPres.PageSetup.SlideHeight
    SldWidth = PPTPres.PageSetup.SlideWidth
    
    'Calculate the middle and center of my slide
    SldCenter = SldWidth / 2
    SldMiddle = SldHeight / 2
    
    'Calculate the center & middle of my shape
    ShpMiddle = PPTShape.Height / 2
    ShpCenter = PPTShape.Width / 2
    
    'Align my shape to the center & middle of my slide.
    With PPTShape
        .Left = SldCenter - ShpCenter
        .Top = SldMiddle - ShpMiddle
    End With

End Sub


Sub ManipulateShapeInPowerPoint_Method4()
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim PPTShapeRng As PowerPoint.ShapeRange
    Dim ShpCnt As Integer
    
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
    
    'Count the number of shapes on my slide
    ShpCnt = PPTSlide.Shapes.Count
    
    'Set a reference to the shape I want to manipulate
    Set PPTShapeRng = PPTSlide.Shapes.Range(Array(ShpCnt))
    
    'Set the dimensions of my shaperange
    With PPTShapeRng
        .Height = 300
        .Width = 300
        .Align msoAlignMiddles, True
        .Align msoAlignCenters, True
    End With
   
End Sub

Sub ManipulateShapeInPowerPoint_Method5()
        
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
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste.Select
    
    'Set the dimensions of the selected shape in my ACTIVE WINDOW
    With PPTApp.ActiveWindow.Selection.ShapeRange
        .Height = 300
        .Width = 300
        .Align msoAlignCenters, True
        .Align msoAlignMiddles, True
    End With

End Sub



















