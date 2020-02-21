import win32com.client

# create a new instance of excel
ExcelApp = win32com.client.GetActiveObject("Excel.Application")

# Let's check to make sure it's visible. If it's not then make it visible
if ExcelApp.Visible == False:
    ExcelApp.Visible = True

# Assuming the Workbook is open, then let's grab the one with that contains the ranges we want to export.
workbook = ExcelApp.Workbooks("RangeBook.xlsm")

# create a dictionary object
RangeDict = {}

'''
    I did something special with this Workbook. I created `Named Ranges`. The benefit of using a `Named Range` is it's
    unique at the Workbook level. That means I don't need to go to each Worksheet and specify the range. I can stay at the 
    workbook level and grab the ranges from there.

    What we have to do to leverage this is use the `Names` collection property that belongs to the Application object. The
    `Names` collection is a list of all the `object names` in the workbook.
'''

# loop through all the named ranges and store them in the dictionary
for namedRng in workbook.Names:
    
    # get the range index so it's position in the Named Ranges Collection.
    rngIndex = namedRng.Index

    # Get the name of the `object` in this case I only have ranges.
    rngName = namedRng.Name

    # I also want a refrence to the the Range itself. Using the `Applciation` object I pass throug the name to the `Range` object.
    rngObj = ExcelApp.Range(namedRng.Name)
    
    # Store it all in a dictionary where the index is the key.
    RangeDict[rngIndex] = {'name':rngName, 'object':rngObj}

print("\nHere is my Range Dicitonary:\n\n{}".format(str(RangeDict)))


# Create a new instance of PowerPoint, USING THE DEFAULT BINDING.
PPTApp = win32com.client.Dispatch("PowerPoint.Application")

# Let's check to make sure it's visible. If it's not then make it visible.
if PPTApp.Visible == False:
    PPTApp.Visible = True

# create a new presentation in the application
PPTPresentation = PPTApp.Presentations.Add()

# loop through each item in our Range dictionary, add a slide, copy the range, and paste it.
for key in RangeDict:
    
    # use the key as the index when creating the slide.
    PPTSlide = PPTPresentation.Slides.Add(Index = key, Layout = 11) # 11 is a ppLayoutTitleOnly (Title Only) Layout.
    
    # copy the range using the value
    RangeDict[key]['object'].Copy()
    
    # paste the range in the slide as a linked OLEObject, it's important to note that what is returned is a ShapeRange Object.
    PPTShapeRange = PPTSlide.Shapes.PasteSpecial(DataType = 10, Link = True) # 10 in a OLEObject.

    print('-'*80)

    # Should return `10` which means msoLinkedOLEObject.
    print("My Shape Range has the Type of: {}".format(PPTShapeRange.Type))

    # How many items are in my Shape Range?
    print("My Shape Range has {} items in it.".format(PPTShapeRange.Count))

    # Here's the odd part, it would appear I have a duplicate Object in it.
    for index in range(1, PPTShapeRange.Count + 1):
        print('')
        print('\t Name:' + PPTShapeRange.Item(index).Name)
        print('\t Type:' + str(PPTShapeRange.Item(index).Type))

    # Set the Height of the Object.
    try:
        # So this fails.
        PPTShapeRange.Height = 200
    except:
        # but this won't
        PPTShapeRange.Item(1).Height = 200

    try:
        # this will also fail
        PPTShapeRange.Width = 200
    except:
        # but this won't
        PPTShapeRange.Item("Object 2").Width = 200
        
    
    # My work around is doing the following. Specify the Shapes Collection and use the Range Method to return a shape range.
    PPTShapeRange = PPTSlide.Shapes.Range(2)

    # Then you can do all the formatting at once.
    PPTShapeRange.Width = 300
    PPTShapeRange.Height = 300
    PPTShapeRange.Align(1, 1) # Align Center, relative to the slide.
    PPTShapeRange.Align(4, 1) # Align Middle, relative to the slide.