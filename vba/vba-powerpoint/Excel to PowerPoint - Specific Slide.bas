Sub ExportMultipleRangesToSpecificSlide()
    
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    
    'Declare Excel Variables
    Dim ExcRng As Range
    Dim RngArray As Variant
    
    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        
    'Create a new presentation
    Set PPTPres = PPTApp.Presentations.Open("C:\Users\305197\Desktop\ShapeFinder.pptm")
    
    'Array that houses all of our ranges that we want to export
    RngArray = Array(Worksheets("One").Range("A1:C5"), Worksheets("One").Range("A10:C14"), Worksheets("Two").Range("C2:E6"))
    
    'Build Slide Array
    SldArray = Array(1, 2, 3)
    
    'Loop through our array, copy the excel range, create a new slide, and paste the range in the slide
    For x = LBound(RngArray) To UBound(RngArray)
    
        'Create a refernce to the range we want to export
        Set ExcRng = RngArray(x)
        
        'Copy the excel range
        ExcRng.Copy
        
        'Create a new slide
        Set PPTSlide = PPTPres.Slides(SldArray(x))
        
        'Paste the range in the slide
        PPTSlide.Shapes.Paste
        
    Next x
    
End Sub
