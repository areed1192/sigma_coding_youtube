import time
import urllib
import requests
from splinter import Browser
from config import password, account_number, client_id

# --------------------- AUTHENTICATION AUTOMATION --------------------------

# define the location of the Chrome Driver - CHANGE THIS!!!!!
executable_path = {'executable_path': r'C:\Users\Alex\Desktop\chromedriver_win32\chromedriver'}

# Create a new instance of the browser, make sure we can see it (Headless = False)
browser = Browser('chrome', **executable_path, headless=False)

# define the components to build a URL
method = 'GET'
url = 'https://auth.tdameritrade.com/auth?'
client_code = client_id + '@AMER.OAUTHAP'
payload = {'response_type':'code', 'redirect_uri':'http://localhost/test', 'client_id':client_code}

# build the URL and store it in a new variable
p = requests.Request(method, url, params=payload).prepare()
myurl = p.url

# go to the URL
browser.visit(myurl)

# define items to fillout form
payload = {'username': account_number,
           'password': password}

# fill out each part of the form and click submit
username = browser.find_by_id("username").first.fill(payload['username'])
password = browser.find_by_id("password").first.fill(payload['password'])
submit   = browser.find_by_id("accept").first.click()

# click the Accept terms button
browser.find_by_id("accept").first.click() 

# give it a second, then grab the url
time.sleep(1)
new_url = browser.url

# grab the part we need, and decode it.
parse_url = urllib.parse.unquote(new_url.split('code=')[1])

# close the browser
browser.quit()



# THE AUTHENTICATION ENDPOINT

# define the endpoint
url = r"https://api.tdameritrade.com/v1/oauth2/token"

# define the headers
headers = {"Content-Type":"application/x-www-form-urlencoded"}

# define the payload
payload = {'grant_type': 'authorization_code', 
           'access_type': 'offline', 
           'code': parse_url, 
           'client_id':client_id, 
           'redirect_uri':'http://localhost/test'}

# post the data to get the token
authReply = requests.post(r'https://api.tdameritrade.com/v1/oauth2/token', headers = headers, data=payload)

# convert it to a dictionary
decoded_content = authReply.json()                       


# grab the access_token
access_token = decoded_content['access_token']
headers = {'Authorization': "Bearer {}".format(access_token)}



# THE DAILY PRICES ENDPOINT 

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/{}/pricehistory".format('GOOG')

# define the payload
payload = {'apikey':client_id,
           'periodType':'day',
           'frequencyType':'minute',
           'frequency':'1',
           'period':'2',
           'endDate':'1556158524000',
           'startDate':'1554535854000',
           'needExtendedHoursData':'true'}

# make a request
content = requests.get(url = endpoint, params = payload)

# convert it dictionary object
data = content.json()



# THE QUOTE ENDPOINT 

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/{}/quotes".format('GOOG')

# define the payload
payload = {'apikey':client_id}

# make a request
content = requests.get(url = endpoint, params = payload)

# convert it dictionary object
data = content.json()

# THE QUOTES ENDPOINT 

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/quotes"

# define the payload
payload = {'apikey':client_id,
           'symbol':'GOOG,MSFT,AAPL'}

# make a request
content = requests.get(url = endpoint, params = payload)

# convert it dictionary object
data = content.json()



# THE MOVERS ENDPOINT 

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/{}/movers".format('$SPX.X')

# define the payload
payload = {'apikey':client_id,
           'direction':'up',
           'change':'value'}

# make a request
content = requests.get(url = endpoint, params = payload)

# convert it dictionary object
data = content.json()



# ACCOUNT ENDPOINT

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/accounts/{}".format('426496649')

# define the payload
payload = {'apikey':client_id}

# make a request
content = requests.get(url = endpoint, headers = headers)

# convert it dictionary object
data = content.json()



# SEARCH INSTURMENTS - CUSIP

# define a CUSIP
CUSIP = '594918104'

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/instruments/{}".format(CUSIP)

# define the payload
payload = {'apikey':client_id}

# make a request
content = requests.get(url = endpoint, params = payload, headers = headers)

# convert it dictionary object
data = content.json()



# SEARCH INSTURMENTS

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/instruments"

# define the payload
payload = {'apikey':client_id,
           'projection':'fundamental',
           'symbol':'MSFT'}

# make a request
content = requests.get(url = endpoint, params = payload, headers = headers)

# convert it dictionary object
data = content.json()



# MARKET HOURS - SINGLE

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/{}/hours".format("FUTURE")

# define the payload
payload = {'apikey':client_id
           'date': '2019-04-29'}

# make a request
content = requests.get(url = endpoint, params = payload, headers = headers)

# convert it dictionary object
data = content.json()



# MARKET HOURS - MULTIPLE

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/marketdata/hours"

# define the payload
payload = {'apikey':client_id,
           'markets':'FUTURE',
           'date': '2019-04-29'}

# make a request
content = requests.get(url = endpoint, params = payload, headers = headers)

# convert it dictionary object
data = content.json()



# ACCOUNTS ENDPOINT

# define an endpoint with a stock of your choice, MUST BE UPPER
endpoint = r"https://api.tdameritrade.com/v1/accounts"

# make a request
content = requests.get(url = endpoint, headers = headers)

# convert it dictionary object
data = content.json()
