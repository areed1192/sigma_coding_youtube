Sub EarlyBinding()

    'Early Binding has several advantages.
    'We can use intellisense to help write our code, this is because the Object model is known.
    'The developer can also compile the code to assure there are no syntax errors.
    'Early binding runs a little faster than late binding.
    
    'Here we are using early binding - We define the object types early on in the code.
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
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutblank)
                
    'Set a reference to the range
    Set ExcRng = Range("A1:C5")
    
    'Copy Range
    ExcRng.Copy

    'Create another slide
    Set PPTSlide = PPTPres.Slides.Add(2, ppLayoutTitleOnly)
        PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=msoTrue

End Sub

Sub LateBinding()

    'Here we are using late binding - We simply define the data type as general objects.
    Dim PPTApp As Object
    Dim PPTPres As Object
    Dim PPTSlide As Object
    Dim ExcRng As Object
    
    'Create a new instance of PowerPoint
    Set PPTApp = CreateObject("PowerPoint.Application")
        PPTApp.Visible = True
    
    'Create a new Presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new Slide
    Set PPTSlide = PPTPres.Slides.Add(1, 1)
                
    'Set a reference to the range
    Set ExcRng = Range("A1:C5")
    
    'Copy Range
    ExcRng.Copy
    
    'Create another slide
    Set PPTSlide = PPTPres.Slides.Add(2, 11)
    
    'Paste the range in the slide
    PPTSlide.Shapes.PasteSpecial DataType:=10
    
End Sub
