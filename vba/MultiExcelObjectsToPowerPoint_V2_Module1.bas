Attribute VB_Name = "Module1"
Option Explicit

Sub ExportMultipleRangesToPowerPoint()
    
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Declare Excel Variables
    Dim ExcObj As Variant
    Dim RngArray, LefArray, TopArray, HgtArray, WidArray As Variant
    Dim x As Integer
    Dim ObjType As Variant
    
    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        
    'Create a new presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Array that houses all of our ranges that we want to export
    RngArray = Array(Sheet2.Range("B2:D5"), Sheet2.ChartObjects(1), Sheet2.ListObjects(1))
    LefArray = Array(59.3, 59.3, 451.5)
    TopArray = Array(270, 34.44, 34.44)
    HgtArray = Array(105.27, 215.25, 131.64)
    WidArray = Array(359.23, 359.25, 449.22)
    
    
    'Loop through our array, copy the excel range, create a new slide, and paste the range in the slide
    For x = LBound(RngArray) To UBound(RngArray)

        'Determine the Object Type
        ObjType = TypeName(RngArray(x))

        'Depending on the Type of object it is, copy it in that manner.
        Select Case ObjType
        
            Case "Range"
                 Set ExcRng = RngArray(x)
                     ExcRng.Copy
                     
            Case "ChartObject"
                Set ExcRng = RngArray(x)
                    ExcRng.Chart.ChartArea.Copy
            
            Case "ListObject"
                Set ExcRng = RngArray(x)
                    ExcRng.Range.Copy
    
        End Select

        'Pause the application
        Application.Wait Now() + #12:00:01 AM#

        
        'Paste the range in the slide
        PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject
        
        Set PPTShape = PPTSlide.Shapes(PPTSlide.Shapes.Count)
            PPTShape.Select
        
        With PPTShape
        
            .Left = LefArray(x)
            .Height = HgtArray(x)
            .Width = WidArray(x)
            .Top = TopArray(x)
        
        End With
        
    Next x
    
End Sub

