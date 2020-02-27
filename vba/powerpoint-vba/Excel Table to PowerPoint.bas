Sub ExportTableToPowerPoint()

    ' OVERVIEW:
    ' This script will create a new PowerPoint Presentation and copy the excel table
    ' we specify to the newly created PowerPoint Presentation.
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Dim Excel Variables
    Dim ExcTbl As ListObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create a new slide in the PowerPoint Presentation
    Set PPTSlide = PPTPres.Slides.Add(1, ppLayoutBlank)
    
    'Create a Table Object variable where specify the sheet the chart is on and the index number of that chart.
    Set ExcTbl = Worksheets("Tables_One").ListObjects(1)
    
        'Copy the Table Object variable we specified above.
        ExcTbl.Range.Copy
   
    'Paste the Chart Object on the Slide that we created above.
    PPTSlide.Shapes.Paste
    

End Sub

Sub ExportMultipleTablesToPowerPoint_Worksheet()

    ' OVERVIEW:
    ' This script will loop through all the List Objects in the active sheet, copy those
    ' list objects and export them to a newly created PowerPoint presentation as a
    ' PowerPoint Shape object.
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Dim Excel Variables
    Dim ExcTbl As ListObject
        
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create an index handler for slide creation.
    SldIndex = 1
    
    'Loop through all the LISTOBJECTS in the ACTIVESHEET.
    For Each ExcTbl In ActiveSheet.ListObjects
    
        'Copy the Excel Table.
        ExcTbl.Range.Copy
   
        'Create a new slide in the Presentation, set the layout to blank, and paste Excel Table on to the newly added slide.
        Set PPTSlide = PPTPres.Slides.Add(SldIndex, ppLayoutBlank)
            PPTSlide.Shapes.Paste
        
        'Increment index so that way we paste the next table on the new slide that is added.
        SldIndex = SldIndex + 1
    
    Next ExcTbl

End Sub

Sub ExportMultipleTablesToPowerPoint_Workbook()

    ' OVERVIEW:
    ' This script will loop through all the List Objects in the active sheet, copy those
    ' list objects and export them to a newly created PowerPoint presentation as a
    ' PowerPoint Shape object.
        
    'Declare PowerPoint Variables
    Dim PPTApp As PowerPoint.Application
    Dim PPTPres As PowerPoint.Presentation
    Dim PPTSlide As PowerPoint.Slide
    Dim PPTShape As PowerPoint.Shape
    
    'Dim Excel Variables
    Dim ExcTbl As ListObject
    Dim WrkSht As Worksheet
            
    'Create a new PowerPoint Application & make it visible.
    Set PPTApp = New PowerPoint.Application
        PPTApp.Visible = True
    
    'Create a new presentation in the PowerPoint Application
    Set PPTPres = PPTApp.Presentations.Add
    
    'Create an index handler for slide creation.
    SldIndex = 1
    
    For Each WrkSht In Worksheets
    
         'Loop through all the LISTOBJECTS in the ACTIVESHEET.
         For Each ExcTbl In ActiveSheet.ListObjects
         
             'Copy the Excel Table.
             ExcTbl.Range.Copy
        
             'Create a new slide in the Presentation, set the layout to blank, and paste Excel Table on to the newly added slide.
             Set PPTSlide = PPTPres.Slides.Add(SldIndex, ppLayoutBlank)
                 PPTSlide.Shapes.Paste
             
             'Increment index so that way we paste the next table on the new slide that is added.
             SldIndex = SldIndex + 1
         
         Next ExcTbl
    
    Next WrkSht

End Sub
