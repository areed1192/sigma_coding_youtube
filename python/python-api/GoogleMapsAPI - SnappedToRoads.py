# isntall the library
# pip install googlemaps

# import the libraries
import googlemaps
from GoogleMapsAPIKey import get_my_key

# Define the API Key.
API_KEY = get_my_key()

# Define the Client
gmaps = googlemaps.Client(key = API_KEY)

# PARAMETERS:
# Path = a single location, or a list of locations, where a location is a string, dict, list, or tuple – The path to be snapped.
# Interpolate =  Whether to interpolate a path to include all points forming the full road-geometry.
#                When true, additional interpolated points will also be returned, resulting in a path 
#                that smoothly follows the geometry of the road, even around corners and through tunnels. 
#                Interpolated paths may contain more points than the original path.
                 
                 # POSSIBLE WAYS TO PASS THROUGH THE PARAMETERS.                 
                 #  '60.170880,24.942795|60.170879,24.942796|60.170877,24.942796'
                 # ['60.170880,24.942795','60.170879,24.942796','60.170877,24.942796']
                 # [('60.170880','24.942795'),('60.170879','24.942796')]
                 # [{"lat": 60.170880, "lng": 24.942795},{"lat": 60.170879, "lng": 24.942796}]

# Get the Snapped Points
road_result  = gmaps.snap_to_roads(path=[('60.170880','24.942795'),('60.170879','24.942796')], interpolate = True)

# Retrieve the individual values from the results.
for snap_point in road_result:

    print(snap_point['location'])
    print(snap_point['originalIndex'])
    print(snap_point['placeId'])

