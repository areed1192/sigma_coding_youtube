# import our libraries
import win32com.client as win32
import pythoncom
import sys

# define our Application Events
class ApplicationEvents:
    
    # define an event inside of our application
    def OnSheetActivate(self, *args):
        print('You Activated a new sheet.')

# define our Workbook Events
class WorkbookEvents:
    
    # define an event inside of our Workbook
    def OnSheetSelectionChange(self, *args):
        
        #print the arguments
        print(args)
        print(args[1].Address)
        args[0].Range('A1').Value = 'You selected cell ' + str(args[1].Address)

# Get the active instance of Excel
xl = win32.GetActiveObject('Excel.Application')

# assign our event to the Excel Application Object
xl_events = win32.WithEvents(xl, ApplicationEvents)

# grab the workbook
xl_workbook = xl.Workbooks('PythonEventsFromExcel.xlsm')

# assign events to Workbook
xl_workbook_events = win32.WithEvents(xl_workbook, WorkbookEvents)

# define initalizer
keepOpen = True

# while there are messages keep displaying them, and also as long as the Excel App is still open
while keepOpen:

    # display the message
    pythoncom.PumpWaitingMessages()

    try:

        # if the workbook count does not equal zero we can assume Excel is open
        if xl.Workbooks.Count != 0:
            keepOpen = True

        # otherwise close the application and exit the script
        else:
            keepOpen = False
            xl = None 
            sys.exit()

    except:

        # if there is an error close excel and exit the script
        keepOpen = False
        xl = None
        sys.exit()
