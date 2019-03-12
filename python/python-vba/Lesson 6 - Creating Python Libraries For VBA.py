# import our libraries
import pythoncom
import win32com.client
import requests
from bs4 import BeautifulSoup
from urllib.request import Request, urlopen

class PythonObjectLibrary:

    # This will create a GUID to register it with Windows, it is unique.
    _reg_clsid_ = pythoncom.CreateGuid()

    # Register the object as an EXE file, the alternative is an DLL file (INPROC_SERVER)
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

    # the program ID, this is the name of the object library that users will use to create the object.
    _reg_progid_ = "Python.APILibrary"

    # this is a description of our object library.
    _reg_desc_ = "This is our Python object library."

    # a list of strings that indicate the public methods for the object. If they aren't listed they are conisdered private.
    _public_methods_ = ['pullUrlLinks', 'ConvertToJson']

    def pullUrlLinks(self, url, ExcRng):   
        '''
           Will take a url, fetch the HTML content, find all the 'a' html tags and extract the links to a python list.
           This list will then be dropped into an Excel range that is specified by the user.
           
           :PARA url: The url that will be fetched by the request.
           :TYPE url: String

           :PARA ExcRng: The range of data where the links will be dropped, the number of cells 
                         in this range will determine how many links come back.
           :TYPE ExcRng: Excel Range Object
        '''

        # Request the URL, and initialize the list that will store the URLs
        req = Request(url)
        html_page = urlopen(req)
        column_list =[]

        # Get the HTML content
        soup = BeautifulSoup(html_page, 'html.parser')
    
        # Loop through each of HREF links and store them in the list.
        for link in soup.findAll('a'):
            row_list = []
            row_list.append(link.get('href'))
            column_list.append(row_list)

        # Get our Worksheet, dispatch the range, and define our worksheet.
        ExcelApp = win32com.client.GetActiveObject("Excel.Application")
        ExcelRng = win32com.client.dynamic.Dispatch(ExcRng)
        WrkSht = ExcelApp.ActiveSheet
        
        # the number of links is determined my the number of cells in our range. Sort of.
        NumOfLinks = ExcelRng.Cells.Count

        # slice the list so we only get the correct number of items.
        if NumOfLinks > len(column_list):
            column_list = column_list
        else:
            column_list = column_list[0:NumOfLinks]        

        ExcelRng.Value = column_list  

    def ConvertToJson(self, KeyRange, DataRange, DestRange):
        '''
           Will take a range of data and convert it into a JSON string object.          
           
           :PARA KeyRange: The range data that the contains the keys needed for the JSON object. It is expected to be horizontal.
           :TYPE KeyRange: Excel Range Object

           :PARA DataRange: The range of data that contains the values for each key. It is expected to be horizontal.
           :TYPE DataRange: Excel Range Object

           :PARA DestRange: A single cell that will be where the JSON string is returned to.
           :TYPE DestRange: Excel Range Object           
        '''

        # get the key values
        KeyRng = win32com.client.dynamic.Dispatch(KeyRange)
        KeyVal = KeyRng.Value[0]
        
        # get the data values
        DataRng = win32com.client.dynamic.Dispatch(DataRange)
        DataVal = DataRng.Value

       # define the destination range.
        DestRng = win32com.client.dynamic.Dispatch(DestRange)

        # initalize the list and dictionary
        dict_list = []

        # loop through each row in the data.
        for row in DataVal:
            json_dict = {}

            # loop through each element in the row
            for index, element in enumerate(row):                
                json_dict[KeyVal[index]] = element

            dict_list.append(json_dict)
        

        DestRng.Value = str(dict_list)

        # release objects from memory.
        KeyRng = None
        DataRng = None
        DestRng = None

if __name__=='__main__':

    print ("Registering COM server...")
    import win32com.server.register
    win32com.server.register.UseCommandLine(PythonObjectLibrary)
