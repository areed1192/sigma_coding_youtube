'''
  THIS CODE WAS PROVIDED BY JIN ZHANG
'''

import time
import urllib
import requests
from splinter import Browser

# Define the Components for Authentication
username = "MY_USERID"
password = "MY_PASSWORD"
api_key = "MY_API_KEY"

# Define path to Web Driver
executable_path = {'executable_path': r'D:\\chromedriver\\chromedriver.exe'}

# Open a new browser
browser = Browser('chrome', **executable_path, headless=False)

# Define the components of request
method = 'GET'
url = 'https://auth.tdameritrade.com/auth?'
client_code = api_key + '@AMER.OAUTHAP'

# Define Payload, MAKE SURE TO HAVE THE CORRECT REDIRECT URI
payload_auth = {'response_type': 'code', 'redirect_uri': 'http://127.0.0.1', 'client_id': client_code}
built_url = requests.Request(method, url, params=payload_auth).prepare()

# Go to the URL
my_url = built_url.url
browser.visit(my_url)

# Fill Out the Form
payload_fill = {'username': username, 'password': password}
browser.find_by_id('username').first.fill(payload_fill['username'])
browser.find_by_id('password').first.fill(payload_fill['password'])
browser.find_by_id('accept').first.click()
time.sleep(1)

# Get the Text Message Box
browser.find_by_text('Can\'t get the text message?').first.click()

# Get the Answer Box
browser.find_by_value("Answer a security question").first.click()

# Answer the Security Questions.
if browser.is_text_present('What is your paternal grandfather\'s first name?'):
    browser.find_by_id('secretquestion').first.fill('myanswer')

elif browser.is_text_present('What was the first name of your first manager?'):
    browser.find_by_id('secretquestion').first.fill('myanswer')

elif browser.is_text_present('What was the name of your first pet?'):
    browser.find_by_id('secretquestion').first.fill('myanswer')

elif browser.is_text_present('What is your father\'s middle name?'):
    browser.find_by_id('secretquestion').first.fill('myanswer')

# Submit results
browser.find_by_id('accept').first.click()

#Trust this device
browser.find_by_xpath('/html/body/form/main/fieldset/div/div[1]/label').first.click()
browser.find_by_id('accept').first.click()

# Sleep and click Accept Terms.
time.sleep(1)
browser.find_by_id('accept').first.click()
