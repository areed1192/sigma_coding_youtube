'REFERNCE: MICROSOFT XML, V6.0
'DOCUMENTATION: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms763742(v=vs.85)

Sub api_requests()

'Declare variables
Dim xml_obj As MSXML2.XMLHTTP60

'Create a reference to the Microsoft XML library
Set xml_obj = New MSXML2.XMLHTTP60

    'Define URL Components
    base_url = "https://maps.googleapis.com/maps/api/place"
    endpoint = "/nearbysearch/xml?"
    
    param_loc = "location="
    param_loc_val = "-33.8670522,151.1957362"
    
    param_radius = "&radius="
    param_radius_val = "1500"
    
    param_type = "&type="
    param_type_val = "restaurant"
    
    param_api = "&key="
    param_api_value = CStr(Worksheets("Sheet2").Range("A1").Value)
    
    'Combine all the different components into a single URL
    api_url = base_url + endpoint + _
              param_loc + param_loc_val + _
              param_radius + param_radius_val + _
              param_type + param_type_val + _
              param_api + param_api_value
                  
    'Open a new request, specify the method and the URL
    xml_obj.Open bstrMethod:="GET", bstrURL:=api_url
    
    'Send the request
    xml_obj.send
    
    'Print the status code, it should be "OK"
    Debug.Print "The Request was " + xml_obj.statusText

    'To parse the info that is sent back, we will store it in a "document" which will leverage a document object model.
    'This model has
    Dim xDoc As MSXML2.DOMDocument60
    Dim xNodes As MSXML2.IXMLDOMNodeList
    Dim xNode As MSXML2.IXMLDOMNode
    
    'First create a new document.
    Set xDoc = New MSXML2.DOMDocument60
    
        'Laod the response text into our document.
        xDoc.LoadXML (xml_obj.responseText)
        
        
    'Find the nodes that contain the results and then select it.
    Set xNodes = xDoc.SelectNodes("/PlaceSearchResponse/result")
    
    'Grab the child nodes of the documents
    Dim xChlNode As IXMLDOMNodeList
    Set xChlNode = xDoc.ChildNodes.Item(1).ChildNodes

    'Print the base name, and node type
    For Each xChl In xChlNode
        Debug.Print xChl.BaseName
        Debug.Print xChl.NodeType 'NODE_ELEMENT (1) - Full List Here: https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms753745(v=vs.85)
    Next
       

        'Grab the name & place id from each of the nodes in the result NodeList.
        For Each xNode In xNodes
            Debug.Print "----------------------------------"
            Debug.Print xNode.SelectSingleNode("name").Text
            Debug.Print xNode.SelectSingleNode("place_id").Text
        Next
        
        'Define the worksheet to house the data
        Dim WrkSht As Worksheet
        Set WrkSht = ThisWorkbook.Worksheets("Sheet1")
        
        
        'Export the data to the worksheet - OPTION ONE
        Count = 1
        For Each xNode In xNodes

            WrkSht.Cells(Count, 1).Value = xNode.SelectSingleNode("name").Text
            WrkSht.Cells(Count, 2).Value = xNode.SelectSingleNode("place_id").Text

            Count = Count + 1

        Next
        
        'Export the data to the worksheet - OPTION TWO
        For i = 0 To xNodes.Length - 1
        
            WrkSht.Cells(i + 1, 4).Value = xNodes.Item(i).SelectSingleNode("name").Text
            WrkSht.Cells(i + 1, 5).Value = xNodes.Item(i).SelectSingleNode("place_id").Text
            
        Next

End Sub
