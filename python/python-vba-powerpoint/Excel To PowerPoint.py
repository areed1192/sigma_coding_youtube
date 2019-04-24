import win32com.client

# create a new instance of excel
ExcelApp = win32com.client.GetActiveObject("Excel.Application")
#ExcelApp.Visible = True

# open the workbook
workbook = ExcelApp.Workbooks.Open(r"C:\Users\Alex\Desktop\RangeBook.xlsm")

# create a dictionary object
RangeDict = {}

# loop through all the named ranges and store them in the dictionary
for namedRng in workbook.Names:
    
    #get the range index & the name
    rngIndex = namedRng.Index
    rngName = namedRng.Name
    
    # set the index as the key & the name as the value
    RangeDict[rngIndex] = rngName

# create a new instance of PowerPoint
PPTApp = win32com.client.Dispatch("PowerPoint.Application")
PPTApp.Visible = True

# create a new presentation in the application
PPTPresentation = PPTApp.Presentations.Add()

# loop through each item in our dictionary, add a slide, copy the range, and paste it.
for key, value in RangeDict.items():
    
    # use the key as the index when creating the slide.
    PPTSlide = PPTPresentation.Slides.Add(Index = key, Layout = 12) # 12 is a blank layout
    
    # copy the range using the value
    ExcelApp.Range(value).Copy()
    
    # paste the range in the slide as a linked OLEObject
    PPTSlide.Shapes.PasteSpecial(DataType = 10, Link = True) # 10 in a OLEObject.

    
# save the presentation.
PPTPresentation.SaveAs(r"C:\Users\Alex\Desktop\ExcelToPowerPoint.pptx")
