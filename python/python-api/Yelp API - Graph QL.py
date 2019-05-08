# ---- pip install gql ----

# import our modules
from gql import gql, Client
from gql.transport.requests import RequestsHTTPTransport
from config import api_key

# define our authentication process.
header = {'Authorization': 'bearer {}'.format(api_key),
          'Content-Type':"application/json"}

# Build the request framework
transport = RequestsHTTPTransport(url='https://api.yelp.com/v3/graphql', headers=header, use_json=True)

# Create the client
client = Client(transport=transport, fetch_schema_from_transport=True)
        
# define a simple query
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

# execute and print this query
print('-'*100)
print(client.execute(query))


# define a simple query - with nested parameters
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
    coordinates {                      
            latitude                      
            longitude                     
        }
    }
}
''')

# execute and print the query
print('-'*100)
print(client.execute(query))


# define a query for multiple businesses - no fragments
query = gql('''{
    b1: business(id: "yelp-san-francisco") {
        name
        id
        rating
        review_count
        photos
    }
    b2: business(id: "garaje-san-francisco") {
        name
        id
        rating
        review_count
        photos
    }
}
''')

# execute and print the query
print('-'*100)
print(client.execute(query))


# define a query for multiple businesses - with fragments
query = gql('''{
    b1: business(id: "yelp-san-francisco") {
        ...basicBizInfo
    }
    b2: business(id: "garaje-san-francisco") {
        ...basicBizInfo
    }
}
fragment basicBizInfo on Business {
    name
    id
    rating
    review_count
    photos
    coordinates {                      
            latitude                      
            longitude                     
        }
}
''')

# execute and print the query
print('-'*100)
print(client.execute(query))


# define a query with a different endpoint, in this case reviews
query = gql('''{
  reviews(business: "garaje-san-francisco") {
    total
    review {
      rating
      text
      user {
        name
        image_url
      }
    }
  }
}
''')

# execute and print the query
print('-'*100)
print(client.execute(query))
