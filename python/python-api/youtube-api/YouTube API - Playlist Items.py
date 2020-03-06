# Import the modules
import requests
import pprint

# Define API KEY
API_KEY = '<MY API KEY>'

# Define our Endpoint
ENDPOINT = 'playlistItems'

# Define Base URL
FINAL_URL = 'https://www.googleapis.com/youtube/v3/{}'.format(ENDPOINT)


# Define search parameters
PARAMETERS = {'part':'snippet,contentDetails,status',
              'maxResults':5,
              'playlistId':'PLcFcktZ0wnNlRhiXWzu7Mkn2GC72f2LXp',
              'key':API_KEY}

# Make a request to the Youtube API
response = requests.get(url = FINAL_URL, params = PARAMETERS)

# Decode our JSON String
youtube_data = response.json()

# Print the response
pprint.pprint(youtube_data)

# Print out Results
for video in youtube_data['items']:
    pprint(video['snippet']['description'])
