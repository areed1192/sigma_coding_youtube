import win32com.client

# Get the Active Instance of PowerPoint
PPTApp = win32com.client.GetActiveObject("PowerPoint.Application")

# Get the Active Presentation
PPTPresentation = PPTApp.ActivePresentation

# Loop through each slide in the Presentation, using the Slides Collection. Slides Collection returns all the Slide Ojbects.
for PPTSld in PPTPresentation.Slides:

    # Loop through each Shape in the Slide, using the Shapes Collection. Shapes Collection returns all the Shape Ojbects.
    for PPTShp in PPTSld.Shapes:

        # check the Shape Type, msoLinkedOLEObject (11) and if it is continue.
        if PPTShp.Type == 11:

            # Grab the LinkFormat Object using the `LinkFormat` property.
            PPTLnkFrmt = PPTShp.LinkFormat

            # Break the Link, using the BreakLink Method.
            PPTLnkFrmt.BreakLink()
