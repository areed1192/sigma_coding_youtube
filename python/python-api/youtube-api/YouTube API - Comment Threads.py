# Import the modules
import requests
import pprint

# Define API KEY
API_KEY = '<MY API KEY>'

# Define Endpoint
ENDPOINT = 'commentThreads'

# Define Base URL
FINAL_URL = 'https://www.googleapis.com/youtube/v3/{}'.format(ENDPOINT)

# Define my parameters of the search
PARAMETERS = {'part':'snippet,replies',
              'allThreadsRelatedToChannelId':'UCxX9wt5FWQUAAz4UrysqK9A',
              'key':API_KEY}

# Make a request to the YouTube API
response = requests.get(url = FINAL_URL, params = PARAMETERS)

# Decode our JSON String
youtube_data = response.json()

# Print the response
pprint.pprint(youtube_data)

# Get the text from a comment
for item in youtube_data['items']:
    print(item['snippet']['topLevelComment']['snippet']['textDisplay'])
