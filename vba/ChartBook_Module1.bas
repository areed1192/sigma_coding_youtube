Attribute VB_Name = "Module1"
Sub ExportChartToPowerPoint()

    ' OVERVIEW:
    ' This script will create a new PowerPoint Presentation and copy the chart
    ' we specify to the newly created PowerPoint Presentation.
        
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
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Charts_One").ChartObjects(1)
        
        'Copy the Chart Object variable we specified above.
        Chrt.Copy
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste
    

End Sub

Sub ExportChartsToPowerPoint_SingleWorksheet()

    ' OVERVIEW:
    ' This script will loop through all the Charts in the Active worksheet
    ' and copy each Chart to a new PowerPoint presentation that we create.
    ' Each chart will get their own individual slide and will be placed in the center of it.
    
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim SldIndex As Integer
        
    'Declare Excel Variables
    Dim Chrt As ChartObject
            
    'Create new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create new presentation in the PowerPoint application.
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create an index handler for slide creation.
    SldIndex = 1
    
    'Loop through all the CHARTOBJECTS in the ACTIVESHEET.
    For Each Chrt In ActiveSheet.ChartObjects
        
        'Copy the Chart
        Chrt.Copy
        
        'Create a new slide in the Presentation, set the layout to blank, and paste chart on to the newly added slide.
        Set PPTSlide = PPTPres.Slides.Add(SldIndex, ppLayoutBlank)
            PPTSlide.Shapes.Paste
        
        'Increment index so that way we paste the next chart on the new slide that is added.
        SldIndex = SldIndex + 1
    
    Next Chrt
    
    
End Sub



Sub ExportChartsToPowerPoint_MultipleWorksheets()

    ' OVERVIEW:
    ' This script will loop through all the worksheets in the Active Workbook
    ' and copy all the Charts to a new PowerPoint presentation that we create.
    ' Each chart will get their own individual slide and will be placed in the center of it.
    
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    Dim SldIndex As Integer
    
    'Declare Excel Variables
    Dim Chrt As ChartObject
    Dim WrkSht As Worksheet
                
    'Create new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create new presentation in the PowerPoint application.
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create an index handler for slide creation.
    SldIndex = 1
    
    'Loop throught all the Worksheets in the Worksheets Collection.
    For Each WrkSht In Worksheets
    
        'Loop through all the CHARTOBJECTS in the ACTIVESHEET.
        For Each Chrt In WrkSht.ChartObjects
            
            'Copy the Chart
            Chrt.Copy
            
            'Create a new slide in the Presentation, set the layout to blank, and paste chart on to the newly added slide.
            Set PPTSlide = PPTPres.Slides.Add(SldIndex, ppLayoutBlank)
                PPTSlide.Shapes.Paste
            
            'Increment index so that way we paste the next chart on the new slide that is added.
            SldIndex = SldIndex + 1
        
        Next Chrt
        
    Next WrkSht
        
End Sub


Sub ExportChartToPowerPoint_PasteMethods()

    ' OVERVIEW:
    ' This script will create a new PowerPoint Presentation and copy the chart
    ' we specify to the newly created PowerPoint Presentation.
        
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
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Charts_One").ChartObjects(1)
    
        'Copy the Chart.
        Chrt.Copy

        'Copy the Chart Area, use when we want to paste as an OLEObject.
        'Chrt.Chart.ChartArea.Copy
    
    'PASTE USING REGULAR PASTE METHOD
    
    'Use Paste method to Paste as Chart Object in PowerPoint
    'PPTSlide.Shapes.Paste
    
    'PASTE USING PASTESPECIAL METHOD
    
    'Paste as Bitmap
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteBitmap
    
    'Paste as Default
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteDefault
    
    'Paste as EnhancedMetafile
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile
    
    'Paste as HTML - DOES NOT WORK WITH CHARTS
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteHTML
    
    'Paste as GIF
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteGIF
    
    'Paste as JPG
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteJPG
    
    'Paste as MetafilePicture
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteMetafilePicture
    
    'Paste as PNG
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPastePNG
    
    'Paste as Shape
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteShape
    
    'Paste as Shape and it is linked.
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteShape, Link:=msoTrue
    
    'Paste as Shape, display it as an icon, change the icon label, and make it a linked icon.
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteShape, DisplayAsIcon:=True, IconLabel:="Link to my Chart", Link:=msoTrue
    
    'Paste as OLEObject and it is linked.
    PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=msoTrue
    
End Sub



