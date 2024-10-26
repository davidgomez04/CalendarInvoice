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

# Get the list of calendars
result = service.calendarList().list().execute()
calendar_id = result['items'][0]['id']

# Define the start and end dates for the current month
now = datetime.now()
start_of_month = datetime(now.year, now.month, 1).isoformat() + 'Z'  # First day of the month
if now.month == 12:
    end_of_month = datetime(now.year + 1, 1, 1) - timedelta(days=1)
else:
    end_of_month = datetime(now.year, now.month + 1, 1) - timedelta(days=1)
end_of_month = end_of_month.isoformat() + 'Z'  # Last day of the month

# Get the events for the current month
events_result = service.events().list(
    calendarId=calendar_id,
    timeMin=start_of_month,
    timeMax=end_of_month,
    singleEvents=True,
    orderBy='startTime'
).execute()

events = events_result.get('items', [])

# Print the events for the current month
if not events:
    print('No events found for this month.')
else:
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))

        # Determine if it's an all-day event
        is_all_day = 'date' in event['start']

        # Format the start and end times
        if is_all_day:
            formatted_start = datetime.fromisoformat(start).strftime('%A, %B %d, %Y (All Day)')
            formatted_end = ''
        else:
            start_datetime = datetime.fromisoformat(start)  # Remove 'Z' and parse
            end_datetime = datetime.fromisoformat(end)
            formatted_start = start_datetime.strftime('%A, %B %d, %Y at %I:%M %p')
            formatted_end = end_datetime.strftime(' to %I:%M %p')

        # Get event details with default values if any are missing
        title = event.get('summary', 'No Title')
        location = event.get('location', 'No Location')
        description = event.get('description', 'No Description')

        # Print the nicely formatted event details
        print(f"Title: {title}")
        print(f"Time: {formatted_start}{formatted_end}")
        print(f"Location: {location}")
        #print(f"Description: {description}")
        print("-" * 40)