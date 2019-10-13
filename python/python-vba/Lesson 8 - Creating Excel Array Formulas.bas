Function LINKSFROMPAGE_List(URL As String, Number_Of_Links As Integer) As Variant

    '
    ' This function takes two arguments, a URL and the number of links you want to return
    ' from the specified URL. The maximum number of links that will be returned is 10.
    ' Additionally, the returned object is list of values organized in a row fashion.
    '
    ' For Example:
    '
    '     My_Links = [Link1, Link2, Link3]
    '
    ' This means in Excel when this formula is called, it will dump the data as followed:
    '
    '     Cell A1 |  Cell B1 |  Cell C1 |
    '     -------------------------------
    '     Link1   |  Link2   |  Link3   |
    '


    'Create Instance of COM Object
    Set PythonWebScraper = VBA.CreateObject("Python.WebScraper")
    
    'Grab the Links, using the "GRABLINKSFROMPAGE_LIST" method.
    Link_Results = PythonWebScraper.GRABLINKSFROMPAGE_List(URL, Number_Of_Links)

    'Return the Results
    LINKSFROMPAGE_List = Link_Results
    
End Function

Function LINKSFROMPAGE_Array(URL As String, Number_Of_Links As Integer) As Variant
    
    '
    ' This function takes two arguments, a URL and the number of links you want to return
    ' from the specified URL. The maximum number of links that will be returned is 10.
    ' Additionally, the returned object is list of values organized in a column fashion.
    '
    ' For Example:
    '
    '     My_Links = [ Link1,
    '                  Link2,
    '                  Link3 ]
    '
    ' This means in Excel when this formula is called, it will dump the data as followed:
    '
    '     Cell A1 |  Link1 |
    '     ------------------
    '     Cell A2 |  Link2 |
    '     ------------------
    '     Cell A3 |  Link3 |

    
    'Create Instance of COM Object
    Set PythonWebScraper = VBA.CreateObject("Python.WebScraper")
    
    'Grab the Links, using the "GRABLINKSFROMPAGE_ARRAY" method.
    Link_Results = PythonWebScraper.GRABLINKSFROMPAGE_ARRAY(URL, Number_Of_Links)

    'Return the Results
    LINKSFROMPAGE_Array = Link_Results

End Function
