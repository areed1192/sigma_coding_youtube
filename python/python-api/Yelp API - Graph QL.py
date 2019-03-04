
# Import Modules
from gql import gql, Client
from gql.transport.requests import RequestsHTTPTransport


# Define API Key, Search Type, and header
MY_API_KEY = '<YOUR API KEY>'
BUSINESS_PATH = 'https://api.yelp.com/v3/graphql'
HEADERS = {'Authorization': 'bearer %s' % MY_API_KEY}

# Build the URL
_transport = RequestsHTTPTransport(
    url='https://api.yelp.com/v3/graphql',
    headers = HEADERS,
    use_json=True
)

# Define the client - The Part the Reaches out to the Server
client = Client(
    transport=_transport,
    fetch_schema_from_transport=True,
)

# Define The Query
query = gql('''
{
  business(id: "garaje-san-francisco") {
    name
    id
    is_claimed
    is_closed
    url
    phone
    display_phone
    review_count
    rating
    photos
  }
}
''')

# Print the Results
print(client.execute(query))
