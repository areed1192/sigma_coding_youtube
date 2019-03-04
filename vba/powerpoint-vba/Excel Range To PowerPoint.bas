Attribute VB_Name = "Module1"
Sub ExportRangeToPowerPoint()

    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    
    Dim ExcRng As Range
    
    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new Presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new Slide
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
                
    'Set a reference to the range
    Set ExcRng = Range("A1:C5")
    
    'Copy Range
    ExcRng.Copy
    
    'Paste the range in the slide
    PPTSlide.Shapes.Paste
    
    'Create another slide
    Set PPTSlide = PPTPres.Slides.Add(2, ppLayoutBlank)
    
    PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=msoTrue

End Sub


Sub ExportMultipleRangeToPowerPoint_Method1()

    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    
    'Declare Excel Variables
    Dim ExcRng As Range
    Dim RngArray As Variant
    Dim ShtArray As Variant
    
    'Populate our arrays
    RngArray = Array("A1:C5", "A10:C14", "C2:E6", "B2:D6")
    ShtArray = Array("One", "One", "Two", "Three")
    
    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new Presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Loop through the range array, create a slide for each range, and copy that range on to the slide.
    For x = LBound(RngArray) To UBound(RngArray)
    
        'Set a reference to the range
        Set ExcRng = Worksheets(ShtArray(x)).Range(RngArray(x))

        'Copy the range
        ExcRng.Copy
        
        'Create a new Slide
        Set PPTSlide = PPTPres.Slides.Add(x + 1, ppLayoutBlank)
        
        'Paste the range in the slide
        PPTSlide.Shapes.Paste
        
    Next x

End Sub


Sub ExportMultipleRangeToPowerPoint_Method2()

    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    
    'Declare Excel Variables
    Dim ExcRng As Range
    Dim RngArray As Variant
    
    'Populate our array
    RngArray = Array(Sheet1.Range("A1:C5"), Sheet1.Range("A10:C14"), Sheet2.Range("C2:E6"), Sheet3.Range("B2:D6"))

    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new Presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Loop through the range array, create a slide for each range, and copy that range on to the slide.
    For x = LBound(RngArray) To UBound(RngArray)
    
        'Set a reference to the range
        Set ExcRng = RngArray(x)
        
        'Copy Range
        ExcRng.Copy
        
        'Enable this line of code if you recieve error about the range not being in the clipboard - This will fix that error by pausing the program for ONE Second.
        'Application.Wait Now + #12:00:01 AM#
        
        'Create a new Slide
        Set PPTSlide = PPTPres.Slides.Add(x + 1, ppLayoutBlank)
        
        'Paste the range in the slide as a linked OLEObject
        PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=msoTrue
    
    Next x

End Sub
