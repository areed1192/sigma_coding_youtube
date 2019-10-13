
# Import our libraries
import pythoncom
import requests
import win32com.client
from bs4 import BeautifulSoup


class PythonObjectLibrary:

    # This will create a GUID to register it with Windows, it is unique and not user friendly ID.
    _reg_clsid_ = pythoncom.CreateGuid()

    # Register the object as an EXE file, the alternative is an DLL file (INPROC_SERVER)
    _reg_clsctx_ = pythoncom.CLSCTX_LOCAL_SERVER

    # The program ID, this is the name of the object library that users will use to create the object. It's also the more User friendly ID
    _reg_progid_ = "Python.WebScraper"

    # This is a description of our object library.
    _reg_desc_ = "This COM Object Server, will allow you to access web scraping functionality found in python."

    # A list of strings that indicate the public methods for the object. If they aren't listed they are conisdered private. Private methods cannot be accessed.
    _public_methods_ = ['Pull_URL_Links', 'GRABLINKSFROMPAGE_ARRAY','GRABLINKSFROMPAGE_LIST']

    def Pull_URL_Links(self, URL_Link, Excel_Dump_Range):   
        '''
            This will serve as an Excel array formula, that takes two inputs (URL_Link and Number_Of_Links) and
            return the number of links from that web page specified in the function.
           
            NAME: URL_Link
            PARA: A valide URL that will redirect to a web page.
            TYPE: String
            REQU: True

            NAME: Excel_Dump_Range
            PARA: An Excel range, that is passed from VBA to Python.
            TYPE: Python COM Object 
            REQU: True   
        '''

        # Request the URL, and store response.
        response = requests.get(URL_Link)

        # Grab the HTML content, and pass it through the parser.
        soup = BeautifulSoup(response.content, 'html.parser')
    
        # Define a list, this will store all the links we want to pass back to Excel.
        returned_links =[]

        # find all the Anchor tags ('a'), that have an HREF (href = True)
        for link in soup.find_all('a', href=True):

            '''

                To make things more complicated, Excel will be expecting an Array back.
                Here is an example of how we need to build this array.

                My_Array = [ [], 
                             [], 
                             [] ]

                Where the "Outer List" serves as a "Column" and the "Inner Lists" serves as "Rows".

            '''

            # First initalize the "Row".
            row_list = []

            # Add the value to the row
            row_list.append(link['href'])

            # Store the row in the "Column", in this case the "Returned Links" list.
            returned_links.append(row_list)

        # Get our Worksheet, dispatch the range, and define our worksheet.
        Excel_Range_to_Dump = win32com.client.dynamic.Dispatch(Excel_Dump_Range)

        # This method will only return the number of links that is equal to the Range Cell Count, assuming there are enough links to return.
        Number_Of_Links = Excel_Range_to_Dump.Cells.Count

        # Make sure to return the right number of links
        if Number_Of_Links > len(returned_links):
            
            # If the number of links to return is greater than the number of links scraped, then return all links.
            returned_links = returned_links

        else:

            # Otherwise, the number of links you scraped is greater then the number of cells in the range, so slice the list.
            returned_links = returned_links[0:Number_Of_Links]        


        Excel_Range_to_Dump.Value = returned_links  

    def GRABLINKSFROMPAGE_ARRAY(self, URL_Link, Number_Of_Links = 1):   
        '''
            This will serve as an Excel array formula, that takes two inputs (URL_Link and Number_Of_Links) and
            return the number of links from that web page specified in the function.
           
            NAME: URL_Link
            PARA: A valide URL that will redirect to a web page.
            TYPE: String
            REQU: True

            NAME: Number_Of_Links
            PARA: The number of links, that would like returned from the web page. Default is 1, max is 10.
                  If value that is passed through is larger than 10, will return only 10.
            TYPE: Integer 
            REQU: True           

        '''

        # This will ensure only 10 are sent back.
        if Number_Of_Links > 10:
            Number_Of_Links = 10

        # Request the URL, and store response.
        response = requests.get(URL_Link)

        # Grab the HTML content, and pass it through the parser.
        soup = BeautifulSoup(response.content, 'html.parser')

        # Define a list, this will store all the links we want to pass back to Excel.
        returned_links =[]

        # find all the Anchor tags ('a'), that have an HREF (href = True)
        for link in soup.find_all('a', href=True):

            '''

                To make things more complicated, Excel will be expecting an Array back.
                Here is an example of how we need to build this array.

                My_Array = [ [], 
                             [], 
                             [] ]

                Where the "Outer List" serves as a "Column" and the "Inner Lists" serves as "Rows".

            '''

            # First initalize the "Row".
            row_list = []

            # Add the value to the row
            row_list.append(link['href'])

            # Store the row in the "Column", in this case the "Returned Links" list.
            returned_links.append(row_list)

        # When you grabbed all the links, return the final list to Excel.
        return returned_links[0:Number_Of_Links]
    
    def GRABLINKSFROMPAGE_LIST(self, URL_Link, Number_Of_Links = 1):   
        '''
            This will serve as an Excel array formula, that takes two inputs (URL_Link and Number_Of_Links) and
            return the number of links from that web page specified in the function.
           
            NAME: URL_Link
            PARA: A valide URL that will redirect to a web page.
            TYPE: String
            REQU: True

            NAME: Number_Of_Links
            PARA: The number of links, that would like returned from the web page. Default is 1, max is 10.
                  If value that is passed through is larger than 10, will return only 10.
            TYPE: Integer 
            REQU: True           

        '''

        # This will ensure only 10 are sent back.
        if Number_Of_Links > 10:
            Number_Of_Links = 10

        # Request the URL, and store response.
        response = requests.get(URL_Link)

        # Grab the HTML content, and pass it through the parser.
        soup = BeautifulSoup(response.content, 'html.parser')

        # Define a list, this will store all the links we want to pass back to Excel.
        returned_links =[]

        # find all the Anchor tags ('a'), that have an HREF (href = True)
        for link in soup.find_all('a', href=True):

            '''

                To make things more complicated, Excel will be expecting an Array back.
                Here is an example of how we need to build this array.

                My_Array = [ [], 
                             [], 
                             [] ]

                Where the "Outer List" serves as a "Column" and the "Inner Lists" serves as "Rows".

            '''
            # Store the HREF in the "Returned Links" list.
            returned_links.append(link['href'])

        # When you grabbed all the links, return the final list to Excel.
        return returned_links[0:Number_Of_Links]


if __name__=='__main__':

    # Let the user know the COM Server is being registered with Windows.
    print ("Registering COM server...")

    # First, import the Register module.
    import win32com.server.register

    # Use the UseCommandLine method, to take your class object and register it with Windows.
    win32com.server.register.UseCommandLine(PythonObjectLibrary)
