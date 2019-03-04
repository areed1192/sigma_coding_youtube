Attribute VB_Name = "Module2"
Option Explicit

Sub CopyMultiObjectsToPPT()

    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Declare Excel Variables
    Dim ExcObj, ObjType, ObjArray As Variant
    Dim LefArray, TopArray, HgtArray, WidArray As Variant
    Dim x As Integer
    
    'Create a new instance of PowerPoint
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
        PPTApp.Activate
        
    'Create a new presentation
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Create array to house objects we want to export
    ObjArray = Array(Sheet2.Range("B2:D5"), Sheet2.ChartObjects(1), Sheet2.ListObjects(1))
    
    'Define my dimension arrays
    LefArray = Array(59.3, 59.3, 451.5)
    TopArray = Array(270, 34.44, 34.44)
    HgtArray = Array(105.27, 215.25, 131.64)
    WidArray = Array(359.23, 359.25, 449.22)
    
    'Loop through the object array and copy each object
    For x = LBound(ObjArray) To UBound(ObjArray)
    
        'Determine Object Type
        ObjType = TypeName(ObjArray(x))
        
        'Depending on the object type, copy it a certain way
        Select Case ObjType
        
            Case "Range"
                Set ExcObj = ObjArray(x)
                    ExcObj.Copy
            
            Case "ChartObject"
                Set ExcObj = ObjArray(x)
                    ExcObj.Chart.ChartArea.Copy
                    
            Case "ListObject"
                Set ExcObj = ObjArray(x)
                    ExcObj.Range.Copy
        
        End Select
        
        'Pause the Excel Application
        Application.Wait Now() + #12:00:01 AM#
        
        'Paste the object in the slide
        PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject
        
        'Set a reference to the shape
        Set PPTShape = PPTSlide.Shapes(PPTSlide.Shapes.Count)
        
        'Set the dimension of my shape
        With PPTShape
            .Left = LefArray(x)
            .Height = HgtArray(x)
            .Width = WidArray(x)
            .Top = TopArray(x)
        End With
    
    Next
    
End Sub








