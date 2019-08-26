Sub VBA_API_JSON()

'Declare variables
Dim xml_obj As MSXML2.XMLHTTP60

'Create a new Request Object
Set xml_obj = New MSXML2.XMLHTTP60

    'Define URL Components
    base_url = "https://maps.googleapis.com/maps/api/place"
    endpoint = "/nearbysearch/json?"
    
    param_loc = "location="
    param_loc_val = "-33.8670522,151.1957362"
    
    param_radius = "&radius="
    param_radius_val = "1500"
    
    param_type = "&type="
    param_type_val = "restaurant"
    
    param_api = "&key="
    param_api_value = CStr(Worksheets("api_key").Range("A1").Value)
    
    'Combine all the different components into a single URL
    api_url = base_url + endpoint + _
              param_loc + param_loc_val + _
              param_radius + param_radius_val + _
              param_type + param_type_val + _
              param_api + param_api_value
    
    
    'Open a new request using our URL
    xml_obj.Open bstrMethod:="GET", bstrURL:=api_url
    
    'Send the request
    xml_obj.send
    
    'Print the status code in case something went wrong
    Debug.Print "The Request was " + CStr(xml_obj.Status)
    
    'Define a few object variables
    Dim Json, JsonPhoto As Object
    Dim result, photo As Dictionary
    
    'Parse the response
    Set Json = JsonConverter.ParseJson(xml_obj.responseText)
    
    'First handle the results section of our response
    For Each result In Json("results")
        
        Debug.Print "--------"
        
        'Grab the ID
        Debug.Print result("id")
        
        'Grab the ID
        Debug.Print result("name")
        
        'Grab the latitude
        Debug.Print result("geometry")("location")("lat")
        
        'Grab the longitude
        Debug.Print result("geometry")("location")("lng")
        
        'Handle the list of photos
        For Each photo In result("photos")
            
            'Grab the height
            Debug.Print photo("height")
            
            'Grab the width
            Debug.Print photo("width")
            
        Next
    Next
    
    'Define the worksheet that will house the data
    Dim WrkSht As Worksheet
    Set WrkSht = ThisWorkbook.Worksheets("json")
    
    'Loop through the parsed data
    Count = 1
    For Each result In Json("results")
    
        'Grab the id
        WrkSht.Cells(Count, 1).Value = result("id")
        
        'Grab the name
        WrkSht.Cells(Count, 2).Value = result("name")
        
        'Grab the id
        WrkSht.Cells(Count, 3).Value = result("geometry")("location")("lat")
        
        'Handle the list of photos
        For Each photo In result("photos")
        
            'Grab the height
            WrkSht.Cells(Count, 4).Value = photo("height")
        
            'Grab the height
            WrkSht.Cells(Count, 5).Value = photo("width")
            
        Next
        
        'Increment the count
        Count = Count + 1
    
    Next
    
    
End Sub
