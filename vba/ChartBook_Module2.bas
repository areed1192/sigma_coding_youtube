Attribute VB_Name = "Module2"
Sub ExportChartToPowerPoint_PasteMethods()
        
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
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Create a Chart Object variable where specify the sheet the chart is on and the index number of that chart.
    Set Chrt = Worksheets("Charts_One").ChartObjects(1)
    
        'Copy the Chart.
        Chrt.Copy
        
        'Chrt.Chart.ChartArea.Copy

    Application.Wait Now + #12:00:01 AM#
   
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
    
    'Paste as Shape, display it as an icon, change the icon label, and make it a linked icon.
    PPTSlide.Shapes.PasteSpecial DataType:=ppPasteShape, DisplayAsIcon:=True, IconLabel:="Link to my Chart", Link:=msoTrue
    
    'Paste as OLEObject and it is linked.
    'PPTSlide.Shapes.PasteSpecial DataType:=ppPasteOLEObject, Link:=msoFalse
    
End Sub
