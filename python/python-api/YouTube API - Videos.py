# Import the modules
import requests
import pprint
import json
from YouTube_API_Key import get_my_api_key

# Define API KEY
DEVELOPER_KEY = get_my_api_key()

# Define Base URL
BASE_URL = 'https://www.googleapis.com/youtube/v3'

# Define Endpoint
ENDPOINT = 'videos'
#ENDPOINT = 'playlistItems'

# Construct URL
final_url = BASE_URL  + '/' + ENDPOINT



# Define my parameters of the search
PARAMETERS = {'part': 'snippet',
              'id': 'qc4yoUqpwEw',
              'key': DEVELOPER_KEY}

# Make a request to the Yelp API
response = requests.get(url = final_url,
                        params = PARAMETERS)

# Decode the response
encoded_response = response.json()

# Get the items section of the response
item_section = encoded_response['items'][0]

# Video ID
print(item_section['id'])

# Channel ID
print(item_section['snippet']['channelId'])

# Get The Tags
print(item_section['snippet']['tags'])

# Get The Tags
print(item_section['snippet']['categoryId'])
