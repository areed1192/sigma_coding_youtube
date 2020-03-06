# Import the modules
import requests
import pprint

# Define API KEY
DEVELOPER_KEY = '<MY API KEY>'

# Define Endpoint
ENDPOINT = 'videos'

# Define Base URL
FINAL_URL = 'https://www.googleapis.com/youtube/v3/{}'.format(ENDPOINT)

# Define my parameters of the search
PARAMETERS = {'part':'snippet',
              'id':'qc4yoUqpwEw',
              'key':DEVELOPER_KEY}

# Make a request to the Youtube API
response = requests.get(url = FINAL_URL, params = PARAMETERS)

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
