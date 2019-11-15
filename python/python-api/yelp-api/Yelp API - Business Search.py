
#Business Search      URL -- 'https://api.yelp.com/v3/businesses/search'
#Business Match       URL -- 'https://api.yelp.com/v3/businesses/matches'
#Phone Search         URL -- 'https://api.yelp.com/v3/businesses/search/phone'

#Business Details     URL -- 'https://api.yelp.com/v3/businesses/{id}'
#Business Reviews     URL -- 'https://api.yelp.com/v3/businesses/{id}/reviews'

#Businesses, Total, Region

# Import the modules
import requests
import json

# Define a business ID
business_id = '4AErMBEoNzbk7Q8g45kKaQ'
unix_time = 1546047836

# Define my API Key, My Endpoint, and My Header
API_KEY = 'YOUR API KEY'
ENDPOINT = 'https://api.yelp.com/v3/businesses/{}/reviews'.format(business_id)
HEADERS = {'Authorization': 'bearer %s' % API_KEY}

# Define my parameters of the search
# BUSINESS SEARCH PARAMETERS - EXAMPLE
#PARAMETERS = {'term': 'food',
#              'limit': 50,
#              'offset': 50,
#              'radius': 10000,
#              'location': 'San Diego'}

# BUSINESS MATCH PARAMETERS - EXAMPLE
#PARAMETERS = {'name': 'Peets Coffee & Tea',
#              'address1': '7845 Highland Village Pl',
#              'city': 'San Diego',
#              'state': 'CA',
#              'country': 'US'}

# Make a request to the Yelp API
response = requests.get(url = ENDPOINT,
                        params = PARAMETERS,
                        headers = HEADERS)

# Conver the JSON String
business_data = response.json()

# print the response
print(json.dumps(business_data, indent = 3))
