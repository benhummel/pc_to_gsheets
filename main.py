###########################



###########################

from __future__ import print_function
from personalcapital import PersonalCapital, RequireTwoFactorException, TwoFactorVerificationModeEnum
from datetime import date, datetime, timedelta
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import getpass
import logging
import os
import pickle
import json


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

SUMMARY_SHEET_NAME = 'wall_chart'
SUMMARY_RANGE = 'A2:B2'

TRANSACTIONS_SHEET_NAME = 'transactions'
TRANSACTIONS_RANGE = 'A2:H'

SUMMARY_SHEET_RANGE = SUMMARY_SHEET_NAME + "_test!" + SUMMARY_RANGE
TRANSACTIONS_SHEET_RANGE = TRANSACTIONS_SHEET_NAME + "_test!" + TRANSACTIONS_RANGE



class PewCapital(PersonalCapital):
    """
    Extends PersonalCapital to save and load session
    So that it doesn't require 2-factor auth every time
    """
    def __init__(self):
        PersonalCapital.__init__(self)
        self.__session_file = 'session.json'

    def load_session(self):
        try:
            with open(self.__session_file) as data_file:    
                cookies = {}
                try:
                    cookies = json.load(data_file)
                except ValueError as err:
                    logging.error(err)
                self.set_session(cookies)
        except IOError as err:
            logging.error(err)

    def save_session(self):
        with open(self.__session_file, 'w') as data_file:
            data_file.write(json.dumps(self.get_session()))

def get_email():
    email = os.getenv('PEW_EMAIL')
    if not email:
        print('You can set the environment variables for PEW_EMAIL and PEW_PASSWORD so the prompts don\'t come up every time')
        return input('Enter email:')
    return email

def get_password():
    password = os.getenv('PEW_PASSWORD')
    if not password:
        return getpass.getpass('Enter password:')
    return password

def import_pc_data():
    email, password = get_email(), get_password()
    pc = PewCapital()
    pc.load_session()

    try:
        pc.login(email, password)
    except RequireTwoFactorException:
        pc.two_factor_challenge(TwoFactorVerificationModeEnum.SMS)
        pc.two_factor_authenticate(TwoFactorVerificationModeEnum.SMS, input('code: '))
        pc.authenticate_password(password)

    accounts_response = pc.fetch('/newaccount/getAccounts')
    
    now = datetime.now()
    date_format = '%Y-%m-%d'
    days = 90
    start_date = '2019-04-01' # (now - (timedelta(days=days+1))).strftime(date_format)
    end_date = (now - (timedelta(days=1))).strftime(date_format)
    transactions_response = pc.fetch('/transaction/getUserTransactions', {
        'sort_cols': 'transactionTime',
        'sort_rev': 'true',
        'page': '0',
        'rows_per_page': '100',
        'startDate': start_date,
        'endDate': end_date,
        'component': 'DATAGRID'
    })
    pc.save_session()

    accounts = accounts_response.json()['spData']
    networth = accounts['networth']
    print(f'Networth: {networth}')

    transactions = transactions_response.json()['spData']
    total_transactions = len(transactions['transactions'])
    print(f'Number of transactions between {start_date} and {end_date}: {total_transactions}')

    summary = {}

    for key in accounts.keys():
        if key == 'networth' or key == 'investmentAccountsTotal':
            summary[key] = accounts[key]

    transactions_output = []  # a list of dicts
    for this_transaction in transactions['transactions']:
        this_transaction_filtered = {
            'date': this_transaction['transactionDate'],
            'account': this_transaction['accountName'],
            'description': this_transaction['description'],
            'category': this_transaction['categoryId'],
            'tags': '',
            'amount': this_transaction['amount'],
            'isIncome': this_transaction['isIncome'],
            'isSpending': this_transaction['isSpending'] 
        }
        transactions_output.append(this_transaction_filtered)

    # print(transactions_output)

    out = [summary, transactions_output]
    return out


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Get PC data
    
    pc_data = import_pc_data()
    summary_data = pc_data[0]
    transaction_data = pc_data[1]

    #     pc_data = {'networth': 486483.22, 'investmentAccountsTotal': 394698.19}
    
    networth = summary_data['networth']
    investments = summary_data['investmentAccountsTotal']


    # reshape transaction data
    # returns a list of lists, where each sub-list is just the transaction values
    print("time to reshape")

    eventual_output = []

    for i in transaction_data:
        this_transaction_list = []
        for key in i.keys():
            this_transaction_list.append(i[key])
        eventual_output.append(this_transaction_list)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=SUMMARY_SHEET_RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print(f"current values: {values}")
        print(f"values to insert: {summary_data}")

        # summary data
        print("uploading summary data...")
        summary_body = {
            "range": SUMMARY_SHEET_RANGE,
              "values": [
                [
                  networth,
                  investments
                ]
              ],
              "majorDimension": "ROWS"
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=SUMMARY_SHEET_RANGE,
            valueInputOption='USER_ENTERED', body=summary_body).execute()
        print(result)

        # transactions data
        print("uploading transactions...")
        transactions_body = {
            "range": TRANSACTIONS_SHEET_RANGE,
              "values": eventual_output,
              "majorDimension": "ROWS"
        }
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=TRANSACTIONS_SHEET_RANGE,
            valueInputOption='USER_ENTERED', body=transactions_body).execute()
        print(result)

if __name__ == '__main__':
    main()
