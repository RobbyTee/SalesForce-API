"""
This program pulls a report from SalesForce, 
identifies actionable items (e.g. schedule an install),
and actions on those items one account at a time.

The primary goals of this script are:
    1. Schedule installation appointments
    2. Prepare and send customized firewall rules
    3. Communicate with the customer and my team members
"""
import os
from automation_library import get_logger, CredentialsManager, SalesForceAutomation, GoogleDriveAutomation, CalCom

# SalesForce Report to pull
report_id = '00O4v000008E412EAC'    # Shipped/Arrived Report

# IVR Type - this must be exactly as the report lists it
project_type = 'VOW Full'

# Firewall Rules
firewall_rules_folder_id = '1l5TLSDFNJpCV22_kfupsAT1XOuZk4dZb'  # Implementation Drive/Automation/Firewall Rules
sheet_id = '1OYVd56jFOnsl0nLd3jVAhuo7z9d553llGFkH4lgdNqI'       # Full Solution


def process_google_doc(pharmacy_name):
    """
    This function pulls information out of the Google Doc.
    Reusability is minimal as this information is unique
    to this script's needs.
    """
    # Set this to the file name that was downloaded
    filename = pharmacy_name + '.txt'

    # Initialize the variables
    local_ip_scheme = None
    opie_ip = "DHCP"
    pms_vendor = None
    contact_work_number = None
    contact_phone_number = None
    it_contact_name = None
    it_contact_email = None

    # Scrape the document
    with open(filename, 'r', encoding='utf-8') as file:
        for line in file:
            if 'IP Address:' in line and 'Pharmacy' not in line and 'Public' not in line:
                local_ip_scheme = line.split(':')[1].strip()
                try:
                    split_ip = local_ip_scheme.split('.')
                    opie_ip = split_ip[0] + '.' + split_ip[1] + '.' + split_ip[2] + '.250'
                except Exception as e:
                    pass
            if 'Pharmacy Software Vendor' in line:
                pms_vendor = line.split(':')[1].strip()
            if 'Primary Work Phone' in line:
                contact_work_number = line.split(':')[1].strip()
            if 'Primary Cell Phone' in line:
                contact_phone_number = line.split(':')[1].strip()
            if 'IT Contact Name' in line:
                it_contact_name = line.split(':')[1].strip()
            if 'IT Contact Email' in line:
                it_contact_email = line.split(':')[1].strip()

    os.remove(filename)

    if len(contact_phone_number) < 7 and len(contact_work_number) > 5:
        contact_phone_number = contact_work_number

    else:
        contact_phone_number = None

    return opie_ip, pms_vendor, contact_phone_number, it_contact_name, it_contact_email


def main():
    # Set up logging
    filename = 'fullsolution.log'
    logger = get_logger(module='main', filename=filename)

    # Load credentials
    manager = CredentialsManager()
    success = manager.load_config()
    if not success:
        print('Failed to load your configuration file. Stopping the script!')
        return

    # Authorize access to Google Account
    google_drive = GoogleDriveAutomation()

    # Verify SalesForce connection
    obtain_access_token = manager.salesforce_access_token()
    if not obtain_access_token:
        print('Failed to get SalesForce access token. Check logs for errors.')
        return

    logger.info('Script successfully initialized.')
    print("Script started successfully!")
    
    # Get the report from Salesforce
    salesforce = SalesForceAutomation()
    success, reason, report = salesforce.get_report(report_id=report_id)
    if not success:
        print(reason)
        return
    
    # This defines the number of rows in the DataFrame
    rows = report.shape[0]

    # The primary functions of this script are held within this For-Loop
    for num,row in enumerate(report, start=0):
        if num >= rows:
            break

        # Grab the useful variables from the report
        account_update = report.iloc[num]['Account Update']
        pharmacy_name = report.iloc[num]['Account Name']
        ivr_type = report.iloc[num]['IVR Type']
        equipment_arrival_date = report.iloc[num]['Equipment Arrival Date']

        # Check if the IVR Type matches the type for the script
        if ivr_type != project_type:
            print(f'\n{pharmacy_name} is type {ivr_type}. Skipping this pharmacy.')
            continue
        
        print(f'\nStarting work on {pharmacy_name}.')
        logger.info(f'Starting work on {pharmacy_name}.')
        
        # Grab all variables from the Account Update
        fields = ['Id', 'Contact_Name__c', 'Contact_Email__c', 
                  'Install_Date_Time__c', 'IVR_Install_Tier__c', 
                  'Customer_Account_Google_URL__c', 'Install_Best_Days__c',
                  'Install_Best_Hours__c', 'Timezone__c', 
                  'Specific_Install_Hours__c', 'Self_Installing__c',
                  'Firewall_Rules_Required__c', 'Contact_Phone__c']
        au = salesforce.get_account_update_info(account_update=account_update,
                                           fields=fields)

        account_update_id = au.get('Id')
        contact_name = au.get('Contact_Name__c')
        contact_email = au.get('Contact_Email__c')
        contact_phone_from_au = au.get('Contact_Phone__c')
        install_date_time = au.get('Install_Date_Time__c')
        install_tier = au.get('IVR_Install_Tier__c')
        install_best_days = au.get('Install_Best_Days__c').split(';')
        install_best_hours = au.get('Install_Best_Hours__c')
        install_specific_hours = au.get('Specific_Install_Hours__c')
        google_url = au.get('Customer_Account_Google_URL__c')
        full_timezone = au.get('Timezone__c')
        self_installing = au.get('Self_Installing__c')
        firewall_rules_required = au.get('Firewall_Rules_Required__c')

        """
        Time to download the Google Doc. The URL could look like these:
        https://drive.google.com/drive/folders/g5hqbc9Nk3MwqJbV7
        https://drive.google.com/drive/folders/g5hqbc9Nk3MwqJbV7?usp=sharing
        """
        id_and_extra = google_url.split(sep='/')[5]
        folder_id = id_and_extra.split('?')[0]
        success = google_drive.download_google_doc(document_name=pharmacy_name,
                                                  drive_folder_id=folder_id)
        if not success:
            print(f"""
*******************************************************
 Did not find a Google Doc named, "{pharmacy_name}".
*******************************************************
    A possible fix for this issue would be:
    1. Open the Google Folder: {google_url}
    2. Rename the Google Doc to the pharmacy name exactly as it appears in SalesForce
    3. Try the script again

Skipping {pharmacy_name} for now and processing the next one.
""")
            continue

        (opie_ip, 
        pms_vendor, 
        contact_phone_number, 
        it_contact_name,
        it_contact_email) = process_google_doc(pharmacy_name=pharmacy_name)

        if not contact_phone_number:
            contact_phone_number = contact_phone_from_au

        # Determine if the pharmacy already has an install date
        if install_date_time:
            print('    O Install is already scheduled')
            logger.info('    O Install is already scheduled')
            payload = {'Status__c': 'Install Requested'}
            salesforce.update_account_update(account_update_id=account_update_id, 
                                             payload=payload)

        # Else schedule install
        else:
            print('    O Must schedule install')
            logger.info('    O Must schedule install')
            
            # Initialize Cal.com portion
            cal = CalCom()

            # Determine install tier
            if install_tier == "Tier 3":
                event_id = 740786
                logger.info('    O Install tier 3')
            elif install_tier == "Tier 2" and not firewall_rules_required:
                event_id = 740772
                logger.info('    O Install tier 2')
            else:
                event_id = 740750
                logger.info('    O Install tier 1')
            
            # Convert TimeZone (e.g. "Eastern Standard Time" to "US/Eastern")
            timezone = cal.convert_timezone(timezone=full_timezone)

            available_slots = cal.get_event_slots(event_id=event_id, 
                                                  start_date=equipment_arrival_date, 
                                                  timezone=timezone)

            # Convert days ("Monday", "Wednesday", etc.) to 
            # dates between this week and next            
            install_best_dates = cal.convert_days_to_dates(preferred_days=install_best_days)
            
            # This is not how I imagined this to work. Needs updating!
            # Picking the Event Slot: Either First Available ...
            if install_best_hours == 'First Available':
                slot = cal.get_first_available(avail_slots=available_slots)
                if slot == None:
                    print('    X No available times in the next week for this pharmacy. ')
                    continue
                event_slot = cal.combine_day_time(day_time=slot, timezone=timezone)
                
            # or specified hours
            else:
                specific_hours = None
                if install_specific_hours:
                    specific_hours = install_specific_hours.split(',')
                
                install_best_times = cal.convert_hours_to_time(preferred_hours=install_best_hours, 
                                                               specific_hours=specific_hours)

                result, slot = cal.compare_pref_to_available(preferred_dates=install_best_dates, 
                                                             preferred_times=install_best_times, 
                                                             available_slots=available_slots)

                result_mapping = {
                    'Perfect Match': '    O There is an available slot that is a perfect match',
                    'Close Enough': '    O Their preferred day is available, but had to compromise on time',
                    'Nothing': f'    X Found no matching slot. Forced this slot: {slot}'
                }

                print(result_mapping[result])

                event_slot = cal.combine_day_time(day_time=slot, timezone=timezone)

            # Book the appointment
            success, reschedule_link = cal.schedule_install(event_id=event_id, 
                                                        event_slot=event_slot,
                                                        pharmacy_name=pharmacy_name, 
                                                        customer_name=contact_name,
                                                        customer_email=contact_email, 
                                                        customer_phone=contact_phone_number,
                                                        timezone=timezone)
            if not success:
                print('    X Ran into an issue with scheduling this pharmacy. See logs')
                
            # Amend the Account Update
            else:
                customers_datetime = salesforce.prepare_install_date(event_slot=event_slot)
                payload = {'Install_Date_Time__c': customers_datetime, 
                            'Status__c': 'Install Requested',
                            'Contact_Phone__c': contact_phone_number,
                            'Reschedule_Install_Appointment__c': reschedule_link}
                
                salesforce.update_account_update(account_update_id=account_update_id, payload=payload)
                print('    O Scheduled install and updated Account Update successfully')
                
                contact_id = salesforce.get_contact_id(account_update_id=account_update_id)
                template_logic = {'ivr_type': ivr_type, 'self install': self_installing}
                success = salesforce.send_email_with_template(template_logic=template_logic,
                                                    contact_id=contact_id,
                                                    account_update_id=account_update_id)
                if not success:
                    print('    X Failed to send appointment confirmation from Account Update')
                else:
                    print('    O Sent appointment confirmation email from Account Update')

                payload = {'Install_Date_Time__c': event_slot}
                salesforce.update_account_update(account_update_id=account_update_id, 
                                                 payload=payload)

        # Send Firewall Rules
        if not firewall_rules_required:
            continue

        else:
            print('    O Must send firewall rules')

            # Get Opie MAC Address
            error, opie_info = salesforce.get_asset_info(account_update=account_update,
                                                         asset_name='Opie',
                                                         fields=['MAC_Address__c'])
            if error:
                opie_mac_address = "None" # The MAC is not necessary for FW Rules
                print(error) 
            else:
                opie_mac_address = opie_info.get('MAC_Address__c')

            # Get PBX Hostname
            error, pbx_info = salesforce.get_asset_info(account_update=account_update,
                                                         asset_name='PBX',
                                                         fields=['Vow_Asset_URL__c'])
            if error:
                print(error) # Hostname is necessary for FW Rules
                print('    X Skipped sending firewall rules')
                continue

            else:
                full_url = pbx_info.get('Vow_Asset_URL__c')
                pbx_hostname = full_url.split('//')[1].split('/')[0]
            
            success, error = google_drive.firewall_rules_spreadsheet(folder_id=firewall_rules_folder_id,
                                                                    sheet_id=sheet_id,
                                                                    pbx_hostname=pbx_hostname,
                                                                    opie_mac_address=opie_mac_address,
                                                                    opie_ip_address=opie_ip,
                                                                    pms_vendor=pms_vendor)
            if not success:
                print(error)
                continue
            
            success, error = google_drive.download_google_sheet(sheet_id=sheet_id,
                                                                destination_file='Firewall Rules.pdf')
            if not success:
                print(error)
                continue
            
            # Email Firewall Rules to Contact, IT Contact, and go-live team
            subject = f'Phone system firewall rules to implement - {pharmacy_name} - [Installation]'
            recipients = [contact_email, it_contact_email, 'ivr.golive@lumistry.com']
            body = """Hello,<br><br>
Please review the attached firewall rules and implement them prior to the installation session.<br><br>

A DHCP pool is required for our phones and integration device. The phones will remain DHCP, but we would like to statically assign the On-Premise Interface Equipment (OPIE). 
Typically, the address at .250 is available on the network. We will statically assign the OPIE to .250 unless you have a conflict.<br><br>

If you have any questions, please reply to this email or call us at (864) 541-0650 and ask for the Installation Team.<br><br>
"""
            success, error = google_drive.email_with_attachement(receiver_emails=recipients,
                                                                 subject=subject,
                                                                 body=body,
                                                                 attachment_path='Firewall Rules.pdf')
            if not success:
                print('    X Failed sending email with firewall rules')
                print(error)
                continue
            
            print('    O Sent firewall rules successfully')

            os.remove('Firewall Rules.pdf')

    input('The script is done. Press ENTER to close this window.')

if __name__ == "__main__":
    main()