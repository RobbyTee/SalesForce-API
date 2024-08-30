# SalesForce-API

## Description
This Python project is for the Lumistry Implementation Team. 


## Features
	1. Schedule install appointment through Cal.com
	2. Send confirmation email to customer from personal or group address
	3. Send firewall rules to customer
	4. Update SalesForce records


## Coming Soon
* Training.py
* Standalone.py
    * Analog
    * PSTN
    * SIP
        * Schedule OPIE install
        * Send SIP appointment link
    * IVR Lite
* Build.py
    * Asset phone extensions in SalesForce
    * Build phones into Vultr PBX
    * Possibly update SalesForce record to “Shipped”?
* Lumistry.py


## Guide to use
	Download this entire project as a .zip
	Extract it to anywhere on your computer
	Ask your project admin to fill in ‘config.json’ and ‘google_auth.json’ with keys
	Open Terminal (Windows: Command Prompt)
	Navigate to the root folder of this project
	Run pip install -r requirements.txt
	Run python setup.py


## Setup.py
This script saves necessary variables used throughout the primary scripts. Variables like your SalesForce username, Security Token, API keys, etc.

### SalesForce Username & Password
Your SalesForce username will be saved; saving your password is optional. Anything the script does in SalesForce, it does so as your user. 

### SalesForce Security Token
1. Open your web browser and log into SalesForce.
2. Click your profile icon in the top right.
3. Click "Settings".
4. Under "My Personal Information" click "Reset My Security Token".
5. Reset your security token.
6. Check your email inbox for an email from SalesForce with your secrity token

### Cal.com API Key
1. Open your web browser and log into Cal.com
2. Bottom left of the screen, click on Settings
3. Under Developer, click on API Keys
4. Click +ADD
5. Name the key anything you want
6. Set it to never expire
7. Click Save
8. Copy the API Key


## Bugs
Report bugs to your project admin. Workarounds for common issues will be posted here until a solution is implemented.
