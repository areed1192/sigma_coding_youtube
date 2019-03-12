 #'playlistId': 'PLcFcktZ0wnNn0VMRzVqV82s4vKpaTii_W',

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
ENDPOINT = 'commentThreads'
#ENDPOINT = 'playlistItems'

# Construct URL
final_url = BASE_URL  + '/' + ENDPOINT

# Define my parameters of the search
PARAMETERS = {'part': 'snippet,replies',
              'allThreadsRelatedToChannelId':'UCxX9wt5FWQUAAz4UrysqK9A',
              'key': DEVELOPER_KEY}

# Make a request to the Yelp API
response = requests.get(url = final_url,
                        params = PARAMETERS)

# Decode the response
encoded_response = response.json()
pprint.pprint(encoded_response)

# Get the text from a comment
for item in encoded_response['items']:
    print(item['snippet']['topLevelComment']['snippet']['textDisplay'])
