Attribute VB_Name = "Module3"
Option Explicit

Public Sub TestCopyPastePic()

'Declare PowerPoint Variables
Dim PPTApp As PowerPoint.Application
Dim PPTSlide As PowerPoint.Slide
Dim PPTPres As PowerPoint.Presentation
Dim PPTShape As PowerPoint.Shape

'Declare Excel Variables
Dim ExcShape As Excel.Shape
Dim WrkSht As Worksheet

    'Check if PowerPoint is active
    On Error Resume Next
       Set PPTApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

    'Open PowerPoint if not active
    If PPTApp Is Nothing Then
       Set PPTApp = New PowerPoint.Application
    End If
    
    'Display the PowerPoint presentation
    PPTApp.Visible = True

    'Create new presentation in PowerPoint
    If PPTApp.Presentations.Count = 0 Then
       PPTApp.Presentations.Add
    End If

    'Create a reference to the Active Presentation
    Set PPTPres = PPTApp.ActivePresentation
    
    'Locate Excel charts to paste into the new PowerPoint presentation
    For Each WrkSht In ActiveWorkbook.Worksheets
    
        'If the Worksheet is visible then continue on.
        If WrkSht.Visible Then
    
        For Each ExcShape In ActiveSheet.Shapes
        
            If ExcShape.Type = msoPicture Then
            
               'Create a new slide
               Set PPTSlide = PPTPres.Slides.Add(PPTPres.Slides.Count + 1, ppLayoutText)
               
               'Go to the new slide
               PPTApp.ActiveWindow.View.GotoSlide PPTPres.Slides.Count
               
               'Copy the Shape
               ExcShape.Copy
               
               'Paste Shape in the slide.
               PPTSlide.Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
               Set PPTShape = PPTSlide.Shapes(PPTSlide.Shapes.Count)
                   PPTShape.Select
       
               'Set the dimensions of the shape.
               With PPTApp.ActiveWindow.Selection.ShapeRange
                    .Left = 25
                    .Top = 150
                    .Width = 250
                    .Left = 500
               End With
               
            End If
            
        Next ExcShape
        
        End If
            
    Next WrkSht

'Activate the PowerPoint App
PPTApp.Activate

'Memory clean up
Set PPTSlide = Nothing
Set PPTApp = Nothing
Set PPTShape = Nothing

End Sub

