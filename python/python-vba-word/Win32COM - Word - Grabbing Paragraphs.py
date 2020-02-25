
import googlemaps
import json
import pprint
import time
import win32com.client as win32

# Grab the Word App
WordApp = win32.gencache.EnsureDispatch("Word.Application")

# Grab the Specific Document.
WordDoc = WordApp.Documents("ApiReport.docx")

# Define the Client
gmaps = googlemaps.Client(key = 'YOUR_API_KEY')

# create a list to help with pagnation
place_result_list = []

# Do a simple nearby search where we specify the location
# in lat/lon format, along with a radius measured in meters
places_result  = gmaps.places_nearby(location='32.715736,-117.161087', 
                                     radius = 40000, 
                                     open_now = False , 
                                     type = 'coffee')

# append the results to the master list.
place_result_list.append(places_result)

# as long as there is a next page token keep making requests.
# make sure to pause it, and append the results to the master list.
while 'next_page_token' in places_result.keys():

    time.sleep(3)
    places_result  = gmaps.places_nearby(page_token = places_result['next_page_token'])
    place_result_list.append(places_result)


places_details_list = []

#loop through each of the places in the results, and get the place details.      
for place in places_result['results']:

    # define the place id, needed to get place details. Formatted as a string.
    my_place_id = place['place_id']

    # define the fields you would liked return. Formatted as a list.
    my_fields = ['name','formatted_phone_number']

    # make a request for the details.
    places_details  = gmaps.place(place_id= my_place_id , fields= my_fields)

    details_list = []
    details_list.append(place['place_id'])
    details_list.append(places_details['result'])
    
    # print the results of the details, returned as a dictionary.
    places_details_list.append(details_list)
    print(places_details)


# Grab the selection range
WrdRng = WordApp.Selection.Range

# Go through the first two items
for place in places_details_list[:2]:
    
    # add a table
    WrdTbl = WordDoc.Tables.Add(WrdRng, 4, 2)

    # define it's style
    WrdTbl.Style = "Grid Table 1 Light - Accent 1"

    # column headers
    WrdTbl.Cell(1, 1).Range.Text = "Attribute"
    WrdTbl.Cell(1, 2).Range.Text = "Value"
    
    for index, detail in enumerate(place):

        # add header
        if index == 0:
            WrdTbl.Cell(index + 2, 1).Range.Text = "Place ID"

            # add value
            WrdTbl.Cell(index + 2, 2).Range.Text = detail
            
        # add header
        elif index == 1:
            WrdTbl.Cell(index + 2, 1).Range.Text = "Phone Number"

            # add value
            WrdTbl.Cell(index + 2, 2).Range.Text = detail['formatted_phone_number']            

            # add header
            WrdTbl.Cell(index + 3, 1).Range.Text = "Name"

            # add value
            WrdTbl.Cell(index + 3, 2).Range.Text = detail['name'] 

    # go to the next selection    
    WrdRng = WordApp.Selection.Next(5).Select
    WrdRng = WordApp.Selection.Range

    print(place)