from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import pickle
import os, json
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
calendar_id = result['items'][1]['id']

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

# Filter and print events that contain "tutor" in the summary
tutor_events = [event for event in events if 'tutor' in event.get('summary', '').lower()]

with open("tutoring_rates.json", "r") as rates_file:
    tutor_rates = json.load(rates_file)

if not tutor_events:
    print('No tutoring events found for this month.')
else:
    for event in tutor_events:
        title = event.get('summary', 'No Title')
        location = event.get('location', 'No Location')
        description = event.get('description', 'No Description')
        tutor_name = title.split("tutor")[0].strip()

        # Get event start and end times
        start_time = event['start'].get('dateTime')
        end_time = event['end'].get('dateTime')

        if start_time and end_time:
            # Calculate event duration in hours
            start_dt = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            end_dt = datetime.fromisoformat(end_time.replace("Z", "+00:00"))
            duration_hours = (end_dt - start_dt).total_seconds() / 3600  # Duration in hours

            try:
                rate = tutor_rates[tutor_name]
                total_price = rate * duration_hours

                # Print event details and calculated price
                print(f"Title: {tutor_name}")
                print(f"Start Time: {start_dt.strftime('%Y-%m-%d %H:%M')}")
                print(f"End Time: {end_dt.strftime('%Y-%m-%d %H:%M')}")
                print(f"Duration: {duration_hours:.2f} hours")
                print(f"Rate: ${rate}")
                print(f"Total Price: ${total_price:.2f}")
            except KeyError:
                print(f"Can't find rate for {tutor_name}")
            print("-" * 40)
