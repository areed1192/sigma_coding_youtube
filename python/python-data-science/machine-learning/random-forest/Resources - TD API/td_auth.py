# Import libraries
import sys
from config import ACCOUNT_NUMBER, ACCOUNT_PASSWORD, CONSUMER_ID, REDIRECT_URI

# Define path to the TD API folder.
path_to_td_folder = r"YOUR_PATH_TO_THE_TD_LIBRARY"

# I'll be needing my TD API Client to get some prices, so I'll need to point a path to it.
sys.path.insert(0, path_to_td_folder)

# import the TDClient, may get an Intellisence error but disregard it.
from td.client import TDClient

# Create a new session
TDSession = TDClient(account_number = ACCOUNT_NUMBER,
                     account_password = ACCOUNT_PASSWORD,
                     consumer_id = CONSUMER_ID,
                     redirect_uri = REDIRECT_URI)

# Login to the session
TDSession.login()