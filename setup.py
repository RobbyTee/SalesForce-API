"""
This script is to be run first before anything else.
Get the preconfigured file, 'config.json' from your admin and
place it into the keys folder. It contains the SalesForce
Connected App's consumer key and consumer secret. The other
fields will be filled in by this script and used throughout
the automation scripts.
"""
import json
import gspread
import gspread
import os
import pytz
import requests
import sys

from datetime import date, datetime
from getpass import getpass
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive
from time import sleep


class CredentialsManager():
    def __init__(self) -> None:
        global config_path
        config_path = "keys/config.json"
        self.config = {}

    
    def load_config(self) -> bool:
        # Load the config file if possible
        try:
            with open(config_path, "r") as file:
                self.config = json.load(file)            
            return True

        except FileNotFoundError:
            return False
    

    def save_config(self) -> None:
        # Save the config file
        with open(config_path, 'w') as file:
            json.dump(self.config, file, indent=4) # Nice formatting


    def user_options(self) -> None:
        print("""
*********************************************************************
* Windows Users: to paste inside of Command Prompt use right-click! *
*********************************************************************
              """)

        # SalesForce Username
        self.config['username'] = input('Enter your SalesForce email address: ')

        # Store password?
        choice = input('Would you like to store your password? y/n: ')
        if choice == 'y':
            self.config['store_password'] = "True"
            self.config['password'] = getpass('Enter your SalesForce password: ')
        else:
            self.config['store_password'] = "False"
        
        # SalesForce Security Token
        print("""\nPlease follow these steps to obtain your SalesForce Security Token:
    1. Open your web browser and log into SalesForce.
    2. Click your profile icon in the top right.
    3. Click "Settings".
    4. Under "My Personal Information" click "Reset My Security Token".
    5. Reset your security token.
    6. Check your email inbox for an email from SalesForce with your secrity token              

For more information: https://help.salesforce.com/s/articleView?id=sf.user_security_token.htm&language=en_US&type=5\n""")
        self.config['security_token'] = getpass('Paste your SalesForce Security Token and press ENTER: ')

        # Input what email to send from (e.g. Firewall Rules)
        self.config['sending_email_address'] = input('\nInput your Lumistry group email address (e.g. ivr.installation@lumistry.com): ')

        # Cal.com API Key
        print("""\nPlease follow these steps to obtain your Cal.com API Key:
    1. Open your web browser and log into Cal.com
    2. Bottom left of the screen, click on Settings
    3. Under Developer, click on API Keys
    4. Click +ADD
    5. Name the key anything you want
    6. Set it to never expire
    7. Click Save
    8. Copy the API Key
              
For less information: https://cal.com/docs/enterprise-features/api/quick-start#get-your-api-keys\n""")
        self.config['cal_com_key'] = getpass('Paste your Cal.com API Key and press ENTER: ')


class GoogleStuff:
    def __init__(self) -> None:
        self.gspread_token_path = 'keys/token_GSpread.json'
        self.pydrive_token_path = 'keys/token_PyDrive.json'
        self.credentials_path = 'keys/google_auth.json'
        self.scope = ["https://www.googleapis.com/auth/spreadsheets",
                        "https://www.googleapis.com/auth/drive.file",
                        "https://www.googleapis.com/auth/drive",
                        "https://www.googleapis.com/auth/gmail.send"]
            

    def authenticate(self):
        # Create token.json for GSpread
        sys.stdout = open(os.devnull, 'w') # Mutes the spam in the terminal
        flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path,
                                                        self.scope)
        creds = flow.run_local_server(port=0)
        with open(self.gspread_token_path, 'w') as token:
            token.write(creds.to_json())

        # Initialize PyDrive2
        gauth = GoogleAuth(settings={
            'client_config_file': self.credentials_path,
            'save_credentials': True,
            'save_credentials_backend': 'file',
            'save_credentials_file': self.pydrive_token_path,
            'get_refresh_token': True
        })

        self.drive = GoogleDrive(gauth)
        sys.stdout = sys.__stdout__ # Unmutes the spam in the terminal

        # This portion is only for testing.
        about = self.drive.GetAbout()
        name = about['user']['displayName']
        email = about['user']['emailAddress']
        return name, email


    def download_google_doc(self, pharmacy_name, drive_folder_id) -> None:
        filename = pharmacy_name + ".txt"

        drive_payload = {'q':  "'" + drive_folder_id + 
                            "' in parents and trashed=false"}
        drive = self.drive.ListFile(drive_payload).GetList()
        file = [file for file in drive if file['title'] == pharmacy_name]
        file_id = file[0]['id']
        mimetype = file[0]['mimeType']
        
        # Check if it's a folder. Open the folder and repeat the search.
        if mimetype == "application/vnd.google-apps.folder":
            self.download_google_doc(pharmacy_name=pharmacy_name,
                                        drive_folder_id=file_id)
            return
        
        # Download the file
        self.drive.CreateFile({'id': file_id}).GetContentFile(filename)
        return


    def update_spreadsheet(self, spreadsheet_id, name, email, date) -> None:
        # Authenticate GSpread
        creds = Credentials.from_authorized_user_file(self.gspread_token_path,
                                                      self.scope)
        client = gspread.authorize(creds)

        # Update the cells with authorizing user's name, email, and date
        sheet = client.open_by_key(spreadsheet_id).get_worksheet(0)
        sheet.update_cell(row=2, col=1, value=name)
        sheet.update_cell(row=2, col=2, value=email)
        sheet.update_cell(row=2, col=3, value=date)

    def test(self):
        try:
            # Grab the authorizing user's name and email address
            name, email = self.authenticate()
            
            """
            Test downloading a Google Doc from the RxWiki domain.
            Anyone with access to the Vow-Implementation drive should
            have success accessing this document.
            """
            self.download_google_doc(pharmacy_name="Test Google Doc", 
                                drive_folder_id="1rxA6h20nRoAkwh_myUt_F6X4SnbQi5-6")
            os.remove("Test Google Doc.txt")

            """
            Test updating a spreadsheet found on the
            Vow-Implementation drive.
            """
            spreadsheet = "10h0VQ7UyiiwD8BAwyapLhIbwv_B6j2Hgktdtreh2708"
            today = str(date.today())
            self.update_spreadsheet(spreadsheet_id=spreadsheet, name=name, 
                                email=email, date=today)
            
            print("Google Drive authenticated successfully")
            return True
        
        except FileNotFoundError:
            sys.stdout = sys.__stdout__ # Unmutes the spam in the terminal
            print("""
File not found: 'keys/google_auth.json'
Please make sure you have a folder called 'keys' in the same folder as this script.
A file called 'google_auth.json' should be inside this 'keys' folder.
                """)
            return False
        
        except:
            return False


class CalCom:
    def __init__(self) -> None:
        self.base_url = 'https://api.cal.com/v1/'
        

    def load_config(self) -> bool:
        global config
        config = {}
        try:
            with open(config_path, 'r') as file:
                config = json.load(file)
                self.api_key = config['cal_com_key']
                return True
        except FileNotFoundError:
            return False
            

    def get_event_slots(self, event_id: int, start_date: str,
                        end_date: str, timezone: str) -> dict:
        payload = {"eventTypeId": event_id, # Integer
                    "startTime": start_date, # DateTime
                    "endTime": end_date, # DateTime
                    "timeZone": timezone} # US/Eastern
        response = requests.get(f'{self.base_url}slots?apiKey={self.api_key}', 
                                params=payload)
        data = response.json()

        # Create a new dictionary to store cleaned data
        cleaned_data = {}

        # Iterate through the original data and restructure it
        for listed_date, times in data['slots'].items():
            cleaned_times = []
            for time_data in times:
                # Extract time using regular expression
                time_match = time_data['time'].split("T")[1].split("-")[0]
                cleaned_times.append(time_match)

            cleaned_data[listed_date] = cleaned_times

        return cleaned_data
    

    def format_datetime(self, day: str, time: str) -> str:
        #Put the start date and start time together in one string
        date_obj = datetime.strptime(day, '%Y-%m-%d')
        time_obj = datetime.strptime(time, '%H:%M:%S').time()
        tz = pytz.timezone('US/Eastern')
        combined_datetime = datetime.combine(date_obj, time_obj)
        combined_datetime = tz.localize(combined_datetime)
        return combined_datetime.strftime('%Y-%m-%dT%H:%M:%S%z')
        

    def create_payload(self, event_type: int, date_time: str) -> dict:
        payload = {'eventTypeId': int, # Cal event type as integer
                    'start': str, # Formatted like: 2024-05-14T08:00:00-04:00
                    'responses': { 
                                'name': str, # Pharmacy name
                                'email': str, # Customer email
                                'City': str,
                                'Attendee': str, # Customer name
                                'location': {
                                    'value': 'phone',
                                    'optionValue': str # Customer cell phone
                                    },
                                'smsReminderNumber': str # Customer cell phone
                                },   
                    'timeZone': str, # Long-form timezone = 'US/Eastern'
                    'language': 'en', 
                    'metadata': {}
                    }
    
        payload['eventTypeId'] = event_type
        payload['start'] = date_time
        payload['responses']['name'] = "TEST APPOINTMENT"
        payload['responses']['email'] = "Python@Test.com"
        payload['responses']['City'] = "Boiling Springs"
        payload['responses']['Attendee'] = "Python Test"
        payload['responses']['location']['optionValue'] = "+18645410650"
        payload['responses']['smsReminderNumber'] = "+18645410650"
        payload['timeZone'] = 'US/Eastern'
        
        return payload


    def schedule_install(self, url: str, payload: dict) -> bool:
        response = requests.post(url, json=payload)
        
        if response.status_code == 200:
            return True, response.json()
        else:
            return False, response.content


    def cancel_appointment(self, booking_id: int) -> bool:
        url = self.base_url + f'bookings/{booking_id}/cancel' + f'?apiKey={self.api_key}'
        response = requests.delete(url)
        if response.status_code == 200:
            return True, response.content
        else:
            return False, response.content
        
    
    def test(self):
        success = self.load_config()
        if not success:
            print('Could not load config.json')
            return
        
        print("""
Which calendar would you like to submit a test for?
    1. Onboarding Review - Marian
    2. Onboarding Review T2 - Makenzy
    3. Phone and IVR Training - Vance
    4. Phone and IVR System Install - Robby/Evan
    5. Phone and Networking T3 - Robby\n
                """)

        # Dictionary to map choices to event IDs
        event_map = {
            1: 740220,  # Marian
            2: 738298,  # Makenzy
            3: 740668,  # Vance
            4: 740750,  # Evan/Robby
            5: 740786   # Robby
        }

        making_choice = True

        while making_choice:
            try:
                choice = int(input("Please make a selection (1-5): "))
                if choice in event_map:
                    event_id = event_map[choice]
                    making_choice = False
                else:
                    print("Invalid choice. Please select a number between 1 and 5.")
            except ValueError:
                print("Please enter a valid integer.")

        today = date.today()
        first_week = today.isocalendar()[1]
        second_week = first_week + 1
        third_week = first_week + 2
        this_year = int(today.strftime("%Y"))
        second_friday = date.fromisocalendar(this_year, second_week, 5)
        third_friday = date.fromisocalendar(this_year, third_week, 5)

        available = self.get_event_slots(event_id=event_id,
                                    start_date=today,
                                    end_date=third_friday,
                                    timezone="US/Eastern")
        
        dates = sorted(available.keys())
        first_date = dates[0]
        first_time = available[first_date][0]
        print(f'Found first available slot on {first_date} at {first_time}')

        formatted_date = self.format_datetime(first_date, first_time)
        payload = self.create_payload(event_type=event_id, date_time=formatted_date)
        url = self.base_url + 'bookings' + f'?apiKey={self.api_key}'
        
        success, response = self.schedule_install(url, payload)
        if not success:
            print(f'Failed: {response}')
            return
        else:
            print('Successfully scheduled a test appointment')
            status, response = self.cancel_appointment(response['id'])
            
            if not status:
                print(f'Failed: {response}')
                return
            else:
                print('Cancelled the test appointment successfully!\n')
                return
        

class SalesForceStuff():  
    def save_config(self) -> None:
        # Save the config file
        with open(config_path, 'w') as file:
            json.dump(config, file, indent=4) # Nice formatting


    def test(self) -> bool:
        if config.get('store_password') == "True":
            password = config.get('password')
        else:
            password = getpass('Input your SalesForce password: ')
        
        response = requests.post(
         'https://login.salesforce.com/services/oauth2/token',
         data={
            'grant_type': 'password',
            'client_id': config['consumer_key'],
            'client_secret': config['consumer_secret'],
            'username': config['username'],
            'password': password + config['security_token']
            }
        )

        if response.status_code == 200:
            token_data = response.json()
            config['access_token'] = token_data['access_token']
            self.save_config()
            print("SalesForce authentication successful.")
            return True
        else:
            return False


def setup():
    manager = CredentialsManager()
    success = manager.load_config()
    if not success:
        print("File not found: 'keys/config.json' does not exist")
        return
    manager.user_options()
    manager.save_config()
    print("Thank you! Tests will now run. . .\n")

    # Authenticates Google and creates tokens
    google = GoogleStuff()
    google_status = google.test()
    if not google_status:
        print("""
*********************************
* Google failed to authenticate *
*********************************
              """)
    
    # Schedules a test appointment on Cal.com; cancels it
    cal = CalCom()
    cal.test()

    # Refresh SalesForce Access Token
    salesforce = SalesForceStuff()
    salesforce_status = salesforce.test()
    if not salesforce_status:
        print("""
*************************************
* SalesForce failed to authenticate *
*************************************
              """)


if __name__ == "__main__":
    setup()
