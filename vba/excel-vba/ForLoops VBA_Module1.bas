Attribute VB_Name = "Module1"
Sub test()

'Create a reference to the Active Presentation
Set PPTPres = newPowerPoint.ActivePresentation

'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
For Each cht In ActiveSheet.ChartObjects

    'Add a new slide where we will paste the chart
    Set activeSlide = PPTPres.Slides.Add(PPTPres.Slides.Count + 1, ppLayoutText)

    'Copy the chart and paste it into the PowerPoint as a Metafile Picture
    cht.Select
    cht.ChartArea.Copy
    Set PPTShape = activeSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture)
        PPTShape.Select
    
    'Set the dimensions, of the pasted chart.
    With newPowerPoint.ActiveWindow.Selection.ShapeRange
        .Left = 15
        .Top = 125
    End With
    
    'Set the dimensions of the text box??
    With activeSlide.Shapes(2)
        .Width = 200
        .Left = 505
    End With

Next

End Sub

Sub CreatePowerPoint()

'First we declare the variables we will be using
Dim newPowerPoint As PowerPoint.Application
Dim PPTPres As PowerPoint.Presentation
Dim PPTSlide As PowerPoint.Slide
Dim ExcCht As Excel.ChartObject

'Look for existing instance
On Error Resume Next
Set newPowerPoint = GetObject(, "PowerPoint.Application")
On Error GoTo 0

'Let's create a new PowerPoint instance, if there is none.
If newPowerPoint Is Nothing Then
   Set newPowerPoint = New PowerPoint.Application
End If

'Make a presentation in PowerPoint, if there is none.
If newPowerPoint.Presentations.Count = 0 Then
   newPowerPoint.Presentations.Add
End If

'Show the PowerPoint
newPowerPoint.Visible = True

'Create a reference to the Active Presentation
Set PPTPres = newPowerPoint.ActivePresentation

'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
For Each ExcCht In ActiveSheet.ChartObjects

    'Create a new slide, and set this as the slide we want to work with.
    Set PPTSlide = PPTPres.Slides.Add(PPTPres.Slides.Count + 1, ppLayoutText)
        PPTSlide.Select
        
    'Copy the chart and paste it into the PowerPoint as a Metafile Picture
    ExcCht.Chart.ChartArea.Copy
    PPTSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
        
    'Set the dimensions, of the pasted chart.
    With newPowerPoint.ActiveWindow.Selection.ShapeRange
        .Left = 15
        .Top = 125
    End With
    
    'Set the dimensions of the text box.
    With PPTSlide.Shapes(2)
        .Width = 200
        .Left = 505
    End With

Next

'Activate the PowerPoint Application.
newPowerPoint.Activate

'Release Objects from Memory
Set PPTSlide = Nothing
Set newPowerPoint = Nothing

End Sub


Sub CreatePowerPoint2()

'Add a reference to the Microsoft PowerPoint Library by:
'1. Go to Tools in the VBA menu
'2. Click on Reference
'3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay

'First we declare the variables we will be using
    Dim newPowerPoint As PowerPoint.Application
    Dim activeSlide As PowerPoint.Slide
    Dim cht As Excel.ChartObject

 'Look for existing instance
    On Error Resume Next
    Set newPowerPoint = GetObject(, "PowerPoint.Application")
    On Error GoTo 0

'Let's create a new PowerPoint
    If newPowerPoint Is Nothing Then
        Set newPowerPoint = New PowerPoint.Application
    End If
'Make a presentation in PowerPoint
    If newPowerPoint.Presentations.Count = 0 Then
        newPowerPoint.Presentations.Add
    End If

'Show the PowerPoint
    newPowerPoint.Visible = True

'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
    For Each cht In ActiveSheet.ChartObjects

    'Add a new slide where we will paste the chart
        newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
        newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
        Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)

    'Copy the chart and paste it into the PowerPoint as a Metafile Picture
        cht.Select
        ActiveChart.ChartArea.Copy
        activeSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select


    'Adjust the positioning of the Chart on Powerpoint Slide
        newPowerPoint.ActiveWindow.Selection.ShapeRange.Left = 15
        newPowerPoint.ActiveWindow.Selection.ShapeRange.Top = 125

        activeSlide.Shapes(2).Width = 200
        activeSlide.Shapes(2).Left = 505

    Next

'AppActivate ("Microsoft PowerPoint")
Set activeSlide = Nothing
Set newPowerPoint = Nothing

End Sub
