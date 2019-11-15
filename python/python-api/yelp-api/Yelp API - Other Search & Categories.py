
#Transaction Search   URL -- 'https://api.yelp.com/v3/transactions/{transaction_type}/search'
#Autocomplete         URL -- 'https://api.yelp.com/v3/autocomplete'

#Categories           URL -- 'https://api.yelp.com/v3/categories'
#Categories Alias     URL -- 'https://api.yelp.com/v3/categories/{alias}'

# Import the modules
import requests
import json

# Define a business ID
category_alias = 'hotdogs'

# Define API Key, Search Type, and header
MY_API_KEY = 'YOUR API KEY'
BUSINESS_PATH = 'https://api.yelp.com/v3/transactions/delivery/search'
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

# Define Phone Parameter - Phone Search - MUST START WITH "+" and the COUNTRY CODE
#PARAMETERS = {'phone': '+18584340001'}

# Define Parameters - Transaction Type
#PARAMETERS = {'location':'San Diego'}

# Define Paramters -  Autocomplete
#PARAMETERS = {'text': 'good food',
#              'latitude': 32.715736,
#              'longitude': -117.161087}
