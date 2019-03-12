# import the libraries
import googlemaps
import json
from GoogleMapsAPIKey import get_my_key

# Define the API Key.
API_KEY = get_my_key()

# Define the Client
gmaps = googlemaps.Client(key = API_KEY)

        # Define Parameters
        # Method One Place ID:  'place_id:ChIJ7-bxRDmr3oARawtVV_lGLtw'
        # Method Two Lat/Lng:   '32.961951,-117.154038'
        # Method Three Address: '7845 Highland Village Pl, San Diego CA, 92129'

place_origin = 'place_id:ChIJ7-bxRDmr3oARawtVV_lGLtw' # This is the Place ID for the San Diego Airport
place_destin = '32.961951,-117.154038' # This is the Geolocation of Peet's Coffee & Tea in San Diego

# Make a reuqest for direction
direction_results = gmaps.directions(origin = place_origin,         # the origin point
                                     destination = place_destin,    # the destination point
                                     mode = 'driving',              # method of transportation
                                     alternatives = True,           # Get more than one possible route
                                     avoid = 'tolls',               # What do we want to avoid
                                     language = 'en-Au',            # The language we want our response in
                                     units = 'metric')              # System of unit measurement
                                                                          
                                     # traffic_model = 'optimistic'
                                     # arrival_time = 1546301024     
                                     # departure_time = 1546301024  
                                     # This setting affects the value returned in the duration_in_traffic
                                     # Can only be used with the departure time parameter.

print(json.dumps(direction_results, indent = 3))
