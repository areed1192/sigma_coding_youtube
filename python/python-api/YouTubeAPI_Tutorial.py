
# import our modules
import requests
from pprint import pprint
from YouTube_API_Key import get_my_api_key

# Define our API Key
API_KEY = get_my_api_key()

# Define our Endpoint
ENDPOINT = 'playlistItems'

# Define Base URL
BASE_URL = 'https://www.googleapis.com/youtube/v3'

# Construct URL
final_url = BASE_URL + '/' + ENDPOINT

# Define search parameters
PARAMETERS = {'part': 'snippet,contentDetails,status',
              'maxResults': 5,
              'playlistId':'PLcFcktZ0wnNlRhiXWzu7Mkn2GC72f2LXp',
              'key': API_KEY}

# Make a request to the YouTube API
response = requests.get(url = final_url, params=PARAMETERS)

# Decode our JSON String
decoded_response = response.json()

# Print out Results

for video in decoded_response['items']:
    pprint(video['snippet']['description'])
