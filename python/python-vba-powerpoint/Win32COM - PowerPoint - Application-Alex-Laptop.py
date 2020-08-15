# import our libraries
import win32com.client

# get the instance of PowerPoint
PPTApp = win32com.client.GetActiveObject("PowerPoint.Application")

# The applicaiton object has different properties about it, so let's explore some of those properties.
print("The Opertating System used By PowerPoint is: {}".format(PPTApp.OperatingSystem))
print("The Path to the PowerPoint Application is: {}".format(PPTApp.Path))
print("The Product Code for PowerPoint is: {}".format(PPTApp.ProductCode))

# grab the Document Windows collection.
DocumentWindows = PPTApp.Windows

PPTApp.Activate()