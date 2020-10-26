# Overview
This script imports data from Personal Capital (personalcapital.com) and exports various components to a Google Sheet. The use case is if you want to regularly take your financial data to a spreadsheet and run additional analytics.

Right now it's hard coded with names of sheets that I'm using. Be sure to change these to your own where indicated. 

Here's an example of what the schema of the Google Sheet should look like:

https://docs.google.com/spreadsheets/d/1DI5oupu-RlZCCzAR023vCgoXkoo_BYU8DmtGTGvExBk/edit#gid=1923416251


# Setup

## Set up environment variables
```
$ export PEW_EMAIL="<your_personal_capital_email>"
$ export PEW_PASSWORD="<your_personal_capital_password>"
```

## Authorize GSheets access
Follow instructions here:
https://developers.google.com/sheets/api/quickstart/python?authuser=1

## Install packages
Preferably in a virtualenv:
```
$ pip install personalcapital
$ pip install requests
$ pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

## Edit variables in `main.py`
You'll need to edit the following variables:
1. SPREADSHEET_ID:  this is the ID of the Google Sheet you'll be reading and writing to. The format as of this writing is:

`https://docs.google.com/spreadsheets/d/<spreadsheet_id>/...`

2. SUMMARY_SHEET_NAME:  within your spreadsheet, the tab that contains your summary data (one row per month, including columns for Net Worth, Investment Portfolio, etc)
3. TRANSACTIONS_SHEET_NAME:  within your spreadsheet, the tab that contains your transactions data (one row per transaction)
4. TRANSACTIONS_START_DATE:  we'll pull transactions starting on this date through yesterday. YYYY-MM-DD format. Recommended start date is whenever you started categorizing your transactions in Personal Capital.


## Run `main.py`

Expected outcome: 
- The `SUMMARY_SHEET_NAME` sheet in Google Sheets will either have the bottom row overwritten with current data, or a new row created if it's not the current month. e.g. if it's currently October and the last row is September, we create a new row.
- The `TRANSACTIONS_SHEET_NAME` sheet in Google Sheets will be completely overwritten with up-to-date values