
# isntall the library
# pip install googlemaps
# pip install Pillow

# import the libraries
import googlemaps
from GoogleMapsAPIKey import get_my_key
from PIL import Image

# Define the API Key.
API_KEY = get_my_key()

# Define the Client
gmaps = googlemaps.Client(key = API_KEY)


def place_search():
    # Do a simple nearby search where we specify the location
    # in lat/lon format, along with a radius measured in meters
    places_result  = gmaps.places_nearby(location='-33.8670522,151.1957362', radius = 10000)

    # loop through each of the places in the results, and get the place details.      
    for place in places_result['results']:

        # define the place id, needed to get place details. Formatted as a string.
        my_place_id = place['place_id']

        # define the fields you would liked return. Formatted as a list.
        my_fields = ['name','formatted_phone_number', 'photo']

        # make a request for the details using the Places API.
        places_details  = gmaps.place(place_id= my_place_id , fields= my_fields)

        # print get the photo id for each photo for each place, returned as a dictionary.
        for photo in places_details['result']['photos']:
        
            # define parameters of our photo request.
            photo_id = photo['photo_reference']
            photo_width = 400                      
            photo_height = 400

            # request the image, using the Places Photot API.
            raw_image_data = gmaps.places_photo(photo_reference = photo_id, max_width = photo_width, max_height = photo_height)

            # raw image data is returned so we will save that raw data to a JPG file.
            f = open('MyDownloadedImage.jpg', 'wb')
        
            # save the raw image data to the file in chunks.
            for chunk in raw_image_data:
                if chunk:
                   f.write(chunk)
            f.close()

            # we will open the newly saved photo, to display the photo.
            im = Image.open('MyDownloadedImage.jpg')
            im.show()

def find_place():

    my_fields = ['name','formatted_address']
    place_results = gmaps.find_place(input = '+18584340001', 
                                     input_type = 'phonenumber',
                                     fields = my_fields)
    print(place_results)



find_place()
