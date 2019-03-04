#Featured Event       URL -- 'https://api.yelp.com/v3/events/featured'
#Event Search         URL -- 'https://api.yelp.com/v3/events'
#Event Lookup         URL -- 'https://api.yelp.com/v3/events/{id}'

# Import the modules
import requests
import json

# Define an event ID
event_id = 'oakland-saucy-oakland-restaurant-pop-up'

# Define API Key, Search Type, and header
MY_API_KEY = 'YOUR API KEY'
BUSINESS_PATH = 'https://api.yelp.com/v3/events'
HEADERS = {'Authorization': 'bearer %s' % MY_API_KEY}

# Define the Parameters of the search
PARAMETERS = {'location':'San Diego'}

# Make a Request to the API, and return results
response = requests.get(url=BUSINESS_PATH, 
                        params=PARAMETERS, 
                        headers=HEADERS)

# Convert response to a JSON String
business_data = response.json()  

# print the data
print(json.dumps(business_data, indent = 3))
