# Import the modules
import requests
import pprint
from YouTube_API_Key import get_my_api_key

# Define API KEY
DEVELOPER_KEY = get_my_api_key()

# Define Base URL
BASE_URL = 'https://www.googleapis.com/youtube/v3/{}'

# Define Endpoint
ENDPOINT = 'search'

# Construct URL
final_url = BASE_URL.format(ENDPOINT)

# Define my parameters of the search
PARAMETERS = {'part': 'snippet',
              'maxResults': 25,
              'q':'Sigma Coding',
              'type':'channel',
              'key': DEVELOPER_KEY}

# Make a request to the Yelp API
response = requests.get(url = final_url,
                        params = PARAMETERS)

# Decode the response
encoded_response = response.json()

pprint.pprint(encoded_response)

# Get the items section of the response
item_section = encoded_response['items'][0]

# Channel ID
print(item_section['snippet']['channelId'])

# Get The Tags
print(item_section['snippet']['channelTitle'])

# Get The Desc.
print(item_section['snippet']['description'])

