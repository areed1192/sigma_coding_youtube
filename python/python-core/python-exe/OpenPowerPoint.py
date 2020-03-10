
import win32com.client

# create a new instance of PowerPoint
PPTApp = win32com.client.Dispatch("PowerPoint.Application")
PPTApp.Visible = True

# create a new presentation in the application
PPTPresentation = PPTApp.Presentations.Add()

# add a new slide to the presentation
PPTPresentation.Slides.Add(Index = 1, Layout = 12)

