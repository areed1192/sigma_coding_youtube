# Import the modules
import requests
import pprint

# Define API KEY
API_KEY = '<MY API KEY>'

# Define Endpoint
ENDPOINT = 'search'

# Define Base URL
FINAL_URL = 'https://www.googleapis.com/youtube/v3/{}'.format(ENDPOINT)

# Define my parameters of the search
PARAMETERS = {'part':'snippet',
              'maxResults':25,
              'q':'Sigma Coding',
              'type':'channel',
              'key':API_KEY}

# Make a request to the Youtube API
response = requests.get(url = FINAL_URL, params = PARAMETERS)

# Decode our JSON String
youtube_data = response.json()

# Print the full response
pprint.pprint(youtube_data)

# Get the items section of the response
item_section = youtube_data['items'][0]

# Channel ID
print(item_section['snippet']['channelId'])

# Get The Tags
print(item_section['snippet']['channelTitle'])

# Get The Desc.
print(item_section['snippet']['description'])

