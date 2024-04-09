# EECS 280 Office Hours Google Sheet To Google Calender Scraper
This script uses the Google Sheets & Google Calendar API to scrape the Office Hours Scheduling spreadsheet and create Google Calender events for a user based on a configured uniqname. Distinct events are created for Main Room office hour sessions and regular breakout room sessions. 

## Setting up the Script: 
1) Clone this repository
2) Create a virtual environment in Python3 ([install if you don't already have it](https://www.python.org/downloads/)): `python3 -m venv env`
3) Activate the virtual environment: `source env/bin/activate`
4) Install dependencies: `pip install -r requirements.txt`
5) Navigate to `config.yml` and fill out the fields specified with TODOs
6) Using your UMich Google account, follow the steps at [Create a Google Cloud Project](https://developers.google.com/workspace/guides/create-project), then the ones outlined at [Enable Google Workspace APIs](https://developers.google.com/workspace/guides/enable-apis) to ensure that the Google Sheets API and Google Calendar API are both enabled. 
7) [Set up your OAuth consent screen](https://developers.google.com/calendar/api/quickstart/python#configure_the_oauth_consent_screen) and [create OAuth Credentials](https://developers.google.com/calendar/api/quickstart/python#authorize_credentials_for_a_desktop_application). Be sure to save the downloaded credentials file
as `credentials.json`, and move it directly inside the project directory. 
8) Run: `python3 scraper.py` . The first time the script is run, you'll need to login with your UMich email and authorize the workflow. After the script is finished running successfully, refreshing your Google Calender should show a new secondary calender named `EECS 280 OH`, with new events. You can run this script regularly, or set up a cron job to sync changes to your OH assignments or the OH Schedule. 

## Setting up the cron-job:
1) Install Google Cloud Command Line tools if you haven't already: [Follow these steps](https://cloud.google.com/sdk/docs/install-sdk)
2) Enable `Google Cloud Build API`. Select `Try Free` if prompted. You may need to insert some billing information, but running this script as a cron job won't exceed the usage limit and result in charges to your Google Account. 

## Limitations:
1) Can't distinguish between online and in-person office hours, as this information isn't readily available in the sheets


## Future Work
1) Make the code less bad
2) Notification configurations?


