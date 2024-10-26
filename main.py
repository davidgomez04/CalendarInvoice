from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os
from datetime import datetime, timedelta

scopes = ['https://www.googleapis.com/auth/calendar']

# Check if token.pkl exists to load existing credentials
if os.path.exists("token.pkl"):
    with open("token.pkl", "rb") as token_file:
        credentials = pickle.load(token_file)
else:
    # If no valid token, create a new one
    flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", scopes=scopes)
    credentials = flow.run_local_server(port=0)
    
    # Save the credentials for next time
    with open("token.pkl", "wb") as token_file:
        pickle.dump(credentials, token_file)

# Build the service
service = build("calendar", "v3", credentials=credentials)