# Import the modules
import requests
import pprint

# Define API KEY
DEVELOPER_KEY = '<MY API KEY>'

# Define Endpoint
ENDPOINT = 'commentThreads'

# Define Base URL
FINAL_URL = 'https://www.googleapis.com/youtube/v3/{}'.format(ENDPOINT)

# Define my parameters of the search
PARAMETERS = {'part':'snippet,replies',
              'allThreadsRelatedToChannelId':'UCxX9wt5FWQUAAz4UrysqK9A',
              'key':DEVELOPER_KEY}

# Make a request to the YouTube API
response = requests.get(url = FINAL_URL, params = PARAMETERS)

# Decode the response
encoded_response = response.json()

# Print the response
pprint.pprint(encoded_response)

# Get the text from a comment
for item in encoded_response['items']:
    print(item['snippet']['topLevelComment']['snippet']['textDisplay'])
