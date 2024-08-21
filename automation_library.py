import base64
import gspread
import json
import logging
import os
import pandas as pd
import pytz
import requests
import sys

from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from getpass import getpass
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from pandas import DataFrame
from pydrive2.auth import GoogleAuth, RefreshError
from pydrive2.drive import GoogleDrive
from simple_salesforce import Salesforce
from time import sleep


def get_logger(module, filename=None):
    log_format = '%(asctime)s  %(name)8s  %(levelname)5s  %(message)s'
    logging.basicConfig(level=logging.INFO,
                        format=log_format,
                        filename=filename)
    return logging.getLogger(module)
        

class CredentialsManager:
    def __init__(self) -> None:
        self.logger = get_logger(module='CredentialsManager')
        self.config_path = 'keys/config.json'        
        
    
    def load_config(self) -> bool:
        # Load the config file if possible
        global config
        config = {}
        try:
            with open(self.config_path, "r") as file:
                config = json.load(file)
            return True
        
        except FileNotFoundError:
            return False
        

    def save_config(self) -> None:
        # Save the config file
        with open(self.config_path, 'w') as file:
            json.dump(config, file, indent=4)


    def salesforce_access_token(self) -> bool:
        """
        Uses your SalesForce account + the SalesForce
        Connected App for this script and gets an
        access token to make calls against the api.
        """
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
            return True
        else:
            self.logger.error(f'Failed to refresh SalesForce access_token '
                            'with code: {response.status_code}')
            self.logger.error(response.content)
            return False


class SalesForceAutomation:
    def __init__(self) -> None:
        self.logger = get_logger(module='SalesForceAutomation')
        self.sf = Salesforce(session_id=config['access_token'], instance_url=config['instance_url'])
        self.headers = {
                  'Authorization': f'Bearer {config["access_token"]}',
                  'Content-Type': 'application/json'
                    }
    

    def get_report(self, report_id: str) -> DataFrame:
      """
      Gets the data from a report and makes a nice table
      to pull data from. Can print the dataframe in your
      terminal to see all the data presented to you!
      """
      report_url = f'{config["instance_url"]}/services/data/v61.0/analytics/reports/{report_id}'
      response = requests.get(report_url, headers=self.headers)

      if response.status_code == 200:
         report_data = response.json()
         rows = report_data['factMap']['0!T']['rows']
         columns = report_data['reportMetadata']['detailColumns']

         # Extract column labels
         column_labels = []
         for column in columns:
            column_info = report_data['reportExtendedMetadata']['detailColumnInfo'].get(column, {})
            column_labels.append(column_info.get('label', column))

         # Extract row data
         data = []
         for row in rows:
            row_data = [cell.get('label', '') for cell in row['dataCells']]
            data.append(row_data)

         # Create DataFrame
         dataframe = pd.DataFrame(data, columns=column_labels)
         if dataframe.empty:
            self.logger.error('The report is empty.')
            self.logger.error(response.content)
            reason = 'The report is empty!'
            return False, reason, None
         else:
            self.logger.info(f'Report ID ({report_id}) obtained successfully!\n {dataframe}')
            return True, None, dataframe
         
      else:
         self.logger.error(f"Failed to get arrived report: {response.status_code}")
         self.logger.error(response.content)
         reason = response.content
         return False, reason, None


    def get_asset_info(self, account_update, asset_name, fields: list) -> dict:
        query = f"SELECT Account__c FROM Account_Update__c WHERE Name = '{account_update}'"
        response = self.sf.query(query)
        account_id = response['records'][0]['Account__c']

        variables = ', '.join(fields)
        query = f"""
                SELECT {variables}
                FROM Asset
                WHERE AccountId = '{account_id}' AND Name LIKE '%{asset_name}%'
                """
        response = self.sf.query(query)
        try:
            result = {field: response['records'][0][field] for field in fields}
        
        except IndexError:
            error = f'    X Could not find {asset_name} in SalesForce. (Likely not assetted)'
            self.logger.error(response)
            return error, None
        
        return False, result


    def get_account_update_info(self, account_update, fields:list) -> dict:
        variables = ', '.join(fields)
        query = f"SELECT {variables} FROM Account_Update__c WHERE Name = '{account_update}'"
        response = self.sf.query(query)
        result = {field: response['records'][0][field] for field in fields}
        return result
    

    def update_account_update(self, account_update_id, payload):
        self.sf.Account_Update__c.update(account_update_id, payload)


    def send_email_with_template(self, template_id, contact_id, account_update_id):
      payload = {
         'inputs': [
                     {'Template_Id': template_id,
                     'Recipient_Id': contact_id,
                     'Account_Update_Id': account_update_id}
                     ]
                  }

      flow_url = f'{config["instance_url"]}/services/data/v61.0/actions/custom/flow/Email_From_Account_Update'
      response = requests.post(flow_url, headers=self.headers, json=payload)

      if response.status_code == 200:
          return True
      else:
          return False


    def get_contact_id(self, account_update_id):
        query = f"SELECT Contact__c FROM Account_Update__c WHERE Id = '{account_update_id}'"
        response = self.sf.query(query)
        if response['totalSize'] > 0:
            return response['records'][0]['Contact__c']
        else:
            return None


class GoogleDriveAutomation:        
    def __init__(self) -> None:
        self.logger = get_logger(module='GoogleDriveAutomation')
        
        """
        Opens a tab in your web browser to authenticate use
        of your Google Account for this python script.
        Creates two tokens for use later in the script.
        """
        self.gspread_token_path = 'keys/token_GSpread.json'
        self.pydrive_token_path = 'keys/token_PyDrive.json'
        self.credentials_path = 'keys/google_auth.json'
        self.scope = ['https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive.file',
                        'https://www.googleapis.com/auth/drive',
                        'https://www.googleapis.com/auth/gmail.send']
        
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
        

    def download_google_doc(self, document_name, drive_folder_id) -> bool:
        filename = document_name + ".txt"

        drive_payload = {'q':  "'" + drive_folder_id + 
                            "' in parents and trashed=false"}
        
        try:
            drive = self.drive.ListFile(drive_payload).GetList()
        except RefreshError:
            os.remove(self.pydrive_token_path)
            print('Please try to run the script again.')
            return

        file = [file for file in drive if file['title'] == document_name]
        try:
            file_id = file[0]['id']
        except IndexError:
            self.logger.error(f'Could not find Google Doc named, "{document_name}".')
            return False
            
        mimetype = file[0]['mimeType']
        
        # Check if it's a folder. Open the folder and repeat the search.
        if mimetype == "application/vnd.google-apps.folder":
            self.download_google_doc(document_name=document_name,
                                        drive_folder_id=file_id)
            return
        
        # Download the file
        self.drive.CreateFile({'id': file_id}).GetContentFile(filename)
        return True
    

    def firewall_rules_spreadsheet(self, folder_id, sheet_id, pbx_hostname, opie_mac_address, opie_ip_address, pms_vendor) -> bool:
        
        spreadsheet_key = {
            'Full Solution': '1OYVd56jFOnsl0nLd3jVAhuo7z9d553llGFkH4lgdNqI'
        }
        
        if sheet_id in spreadsheet_key.values():
            ivr_type = [key for key, value in spreadsheet_key.items() if value == sheet_id][0]
        else:
            error = '    X Invalid spreadsheet ID. (Firewall rules skipped!)'
            return False, error
        
        # Authenticate GSpread
        creds = Credentials.from_authorized_user_file(self.gspread_token_path,
                                                      self.scope)
        client = gspread.authorize(creds)

        pms_server_ip = pms_vendor + " Server IP Address"

        if ivr_type == 'Full Solution':
            sheet = client.open_by_key(sheet_id).get_worksheet(0)
            sheet.update_cell(row=29, col=2, value=pbx_hostname)
            sheet.update_cell(row=31, col=2, value=opie_mac_address)
            sheet.update_cell(row=32, col=2, value=opie_ip_address)
            sheet.update_cell(row=34, col=2, value=pms_server_ip)
            sleep(2) # Gives the spreadsheet time to convert the hostname to an IP
            return True, None


    def download_google_sheet(self, sheet_id, destination_file):
        try:
            file = self.drive.CreateFile({'id': sheet_id})

            file.GetContentFile(destination_file, mimetype='application/pdf')

            return True, None
        except Exception as error:
            return False, error
    
    
    def email_with_attachement(self, sender_email, receiver_emails, subject, body, attachment_path):
        try:
            # Authenticate with Gmail API
            creds = Credentials.from_authorized_user_file(self.gspread_token_path, self.scope)
            service = build('gmail', 'v1', credentials=creds)
            
            # Create recipients string
            recipients = ', '.join(receiver_emails)

            # Create the email
            message = MIMEMultipart('mixed')
            message['to'] = recipients
            message['from'] = sender_email
            message['subject'] = subject

            # Create alternative HTML part for the body with inline image
            msg_related = MIMEMultipart('related')
            msg_alternative = MIMEMultipart('alternative')
            msg_related.attach(msg_alternative)
            message.attach(msg_related)

            # Attach the body with signature, referencing the image by its CID
            gmail_signature = f"""
            <br><br>
            <table>
                <tr>
                    <td>
                        <img src="cid:signature_image" alt="Signature Image" width="100">
                    </td>
                    <td style="padding-left: 10px;">
                        <b>IVR Implementation Team</b><br>
                        P: (864) 541-0650<br>
                        <a href="http://lumistry.com">Lumistry.com</a><br>
                    </td>
                </tr>
            </table>
            """
            body_with_signature = f"{body}{gmail_signature}"
            msg_alternative.attach(MIMEText(body_with_signature, 'html'))

            # Attach the image as inline content
            with open('resources/Lumistry.png', 'rb') as img_file:
                mime_image = MIMEImage(img_file.read())
                mime_image.add_header('Content-ID', '<signature_image>')
                mime_image.add_header('Content-Disposition', 'inline', filename="signature_image.png")
                msg_related.attach(mime_image)

            if attachment_path:
                with open(attachment_path, 'rb') as attachment:
                    mime_base = MIMEBase('application', 'octet-stream')
                    mime_base.set_payload(attachment.read())
                    encoders.encode_base64(mime_base)
                    mime_base.add_header('Content-Disposition', f'attachment; filename={attachment_path}')
                    message.attach(mime_base)

            # Encode the message as a base64url encoded string
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

            # Create the message object
            message_object = {'raw': raw_message}

            # Send the email
            service.users().messages().send(userId="me", body=message_object).execute()
            return True, None

        except Exception as error:
            return False, error
        

class CalCom:
    def __init__(self) -> None:
        self.logger = get_logger(module='CalCom')
        self.config_path = 'keys/config.json'
        self.base_url = 'https://api.cal.com/v1/'

        self.today = date.today()
        self.this_week = self.today.isocalendar()[1]
        self.next_week = self.this_week + 1
        self.this_year = int(self.today.strftime("%Y"))
        self.next_friday = date.fromisocalendar(self.this_year, self.next_week, 5)

        with open(self.config_path, 'r') as file:
            config = json.load(file)
            self.api_key = config['cal_com_key']

        self.timezone_mapping = {
            "Eastern Standard Time": "US/Eastern",
            "Eastern Daylight Time": "US/Eastern",
            "Central Standard Time": "US/Central",
            "Central Daylight Time": "US/Central",
            "Mountain Standard Time": "US/Mountain",
            "Mountain Daylight Time": "US/Mountain",
            "Pacific Standard Time": "US/Pacific",
            "Pacific Daylight Time": "US/Pacific",
            "Alaska Standard Time": "US/Alaska",
            "Hawaii-Aleutian Standard Time": "US/Hawaii",
            "Hawaii-Aleutian Daylight Time": "US/Hawaii",
            "Atlantic Standard Time": "America/Puerto_Rico",
            "Guam Standard Time": "Pacific/Guam"
        }

        self.hours_mapping = {
            'Morning': ['08:00:00', '09:00:00', '10:00:00', '11:00:00'],
            'Afternoon': ['12:00:00', '13:00:00', '14:00:00', '15:00:00', '16:00:00']
        }

        self.days_mapping = {
            'Monday': 1,
            'Tuesday': 2,
            'Wednesday': 3,
            'Thursday': 4,
            'Friday': 5
        }

    def get_event_slots(self, event_id: int, start_date: datetime, timezone: str) -> dict:
        payload = {"eventTypeId": event_id, # Integer
                    "startTime": start_date, # DateTime
                    "endTime": self.next_friday, # DateTime
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
        self.logger.info(f'Available times to install in {timezone}: {cleaned_data}')
        return cleaned_data
        
    
    def convert_timezone(self, timezone: str):
        iana_timezone = self.timezone_mapping.get(timezone)
        if iana_timezone:
            return iana_timezone
        else:
            self.logger.error('Timezone did not convert successfully')
            return None


    def convert_days_to_dates(self, preferred_days: list) -> list:
        dates = []
        for name, value in self.days_mapping.items():
            if name in preferred_days:
                dates.append(date.fromisocalendar(self.this_year, self.next_week, value).strftime('%Y-%m-%d'))
                test = date.fromisocalendar(self.this_year, self.this_week, value)

                if self.today < test:
                    dates.append(test.strftime('%Y-%m-%d'))
        
        dates.sort()
        return dates
    

    def convert_hours_to_time(self, preferred_hours, specific_hours=None) -> list:
        """
        If Specifc Hours are listed, they are prioritized. Otherwise, take
        the general term "Morning" or "Afternoon" and convert it against the
        map above.
        """
        if specific_hours:
            times = []
            for range_str in specific_hours:
                if '-' in range_str:
                    start, end = map(int, range_str.split('-'))
                    # Convert to 24-hour format if less than 6
                    start = start + 12 if start < 6 else start
                    end = end + 12 if end < 6 else end
                    # Handle range crossing midnight (if needed, adjust as required)
                    if end < start:
                        end += 12
                    times.extend([f"{hour:02d}:00:00" for hour in range(start, end + 1)])
                else:
                    hour = int(range_str)
                    # Convert to 24-hour format if less than 6
                    hour = hour + 12 if hour < 6 else hour
                    times.append(f"{hour:02d}:00:00")
    
            return sorted(set(times))  # Remove duplicates and sort the list
        
        elif preferred_hours in self.hours_mapping:
            return self.hours_mapping[preferred_hours]
        

    def get_first_available(self, avail_slots: dict) -> str:
        for day, times in avail_slots.items():
            for time in times:
                if time >= '08:00:00':
                    return {'day': day, 'time': time}
        
        self.logger.critical('No available install slots were found in the next week!')
        return 


    def compare_pref_to_available(self, preferred_dates: list, preferred_times: list, 
                                  available_slots: dict) -> dict:
        self.logger.info(f'Customer\'s preferred dates: {preferred_dates}')
        self.logger.info(f'Customer\'s preferred times: {preferred_times}')
        
        for day in preferred_dates:
            if day not in available_slots:
                continue  # Skip if the preferred day is not available
            
            for pref_time in preferred_times:
                if pref_time in available_slots[day]:  # Perfect match
                    self.logger.info(f'Found a perfect match on {day} at {pref_time}')
                    return 'Perfect Match', {'day': day, 'time': pref_time}

                # Get within 2 hours of their preference
                pref_time_obj = datetime.strptime(pref_time, '%H:%M:%S')
                closest_time = None
                min_difference = float('inf')

                for avail_time in available_slots[day]:
                    avail_time_obj = datetime.strptime(avail_time, '%H:%M:%S')
                    difference = (avail_time_obj - pref_time_obj).total_seconds()

                    if 0 < difference <= 7200 and difference < min_difference:  # within 2 hours
                        closest_time = avail_time
                        min_difference = difference
                
                if closest_time:
                    self.logger.info(f'Matched customer\'s preferred day ({day}), but compromised on time ({closest_time})')
                    return 'Close Enough', {'day': day, 'time': closest_time}
        
        self.logger.warning('Did not match their preferred time to an available slot. Scheduling first available slot!')
        first_available = self.get_first_available(available_slots)
        return 'Nothing', first_available
    

    def combine_day_time(self, day_time: dict, timezone) -> str:
        #Put the start date and start time together in one string
        date_obj = datetime.strptime(day_time.get('day'), '%Y-%m-%d')
        time_obj = datetime.strptime(day_time.get('time'), '%H:%M:%S').time()
        tz = pytz.timezone(timezone)
        combined_datetime = datetime.combine(date_obj, time_obj)
        combined_datetime = tz.localize(combined_datetime)
        return combined_datetime.strftime('%Y-%m-%dT%H:%M:%S%z')
    

    def schedule_install(self, event_id, event_slot, pharmacy_name, customer_name, customer_email, customer_phone, timezone):
        
        payload = {'eventTypeId': event_id,     # Cal event type as integer
                    'start': event_slot,        # Formatted like: 2024-05-14T08:00:00-04:00
                    'responses': { 
                                'name': pharmacy_name,
                                'email': customer_email,
                                'Attendee': customer_name,
                                'smsReminderNumber': f'+1{customer_phone}',
                                'location': {
                                    'value': 'phone',
                                    'optionValue': f'+1{customer_phone}'
                                    }
                                },   
                    'timeZone': timezone,    # 'US/Eastern'
                    'language': 'en', 
                    'metadata': {}
                    }

        url = self.base_url + 'bookings?apiKey=' + self.api_key
        response = requests.post(url, json=payload)

        if response.status_code == 200:
            response_json = response.json()
            uid = response_json.get('uid')
            reschedule_link = 'https://cal.com/reschedule/' + uid
            return True, reschedule_link
        else:
            self.logger.error(f'Status {response.status_code} when scheduling install.')
            self.logger.error(response.content)
            return False, None
        

    def convert_to_eastern_time(self, date_string):
        # Parse the input datetime string
        original_format = "%Y-%m-%dT%H:%M:%S%z"  # Assuming the input format includes timezone info
        dt = datetime.strptime(date_string, original_format)
        
        # Check if the input has timezone info
        if dt.tzinfo is None:
            raise ValueError("The datetime string must include timezone information.")
        
        # Define the target timezone
        eastern_tz = pytz.timezone('US/Eastern')
        
        # Convert to the target timezone
        # First convert to UTC, then to the target timezone
        utc_dt = dt.astimezone(pytz.utc)
        eastern_dt = utc_dt.astimezone(eastern_tz)
        
        # Return the formatted string in Eastern Time
        return eastern_dt.strftime(original_format)
