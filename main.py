###########################
# 
# Personal Capital to Google Sheets
# Ben Hummel, 2020
# 
###########################

from __future__ import print_function
from personalcapital import PersonalCapital, RequireTwoFactorException, TwoFactorVerificationModeEnum
from datetime import date, datetime, timedelta
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import logging
import os
import pickle
import json
import getpass


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

SUMMARY_SHEET_NAME = 'wall_chart'
TRANSACTIONS_SHEET_NAME = 'transactions'

TRANSACTIONS_START_DATE = '2019-04-01' # the date to start pulling transactions from. 
TRANSACTIONS_END_DATE = (datetime.now() - (timedelta(days=1))).strftime('%Y-%m-%d')

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
		pc.two_factor_authenticate(TwoFactorVerificationModeEnum.SMS, input('Enter 2-factor code: '))
		pc.authenticate_password(password)

	accounts_response = pc.fetch('/newaccount/getAccounts')
	
	transactions_response = pc.fetch('/transaction/getUserTransactions', {
		'sort_cols': 'transactionTime',
		'sort_rev': 'true',
		'page': '0',
		'rows_per_page': '100',
		'startDate': TRANSACTIONS_START_DATE,
		'endDate': TRANSACTIONS_END_DATE,
		'component': 'DATAGRID'
	})
	pc.save_session()
	accounts = accounts_response.json()['spData']
	networth = accounts['networth']
	print(f'Networth: {networth}')

	transactions = transactions_response.json()['spData']
	total_transactions = len(transactions['transactions'])
	print(f'Number of transactions between {TRANSACTIONS_START_DATE} and {TRANSACTIONS_END_DATE}: {total_transactions}')

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
			'amount': this_transaction['amount'], # always a positive int
			'isIncome': this_transaction['isIncome'],
			'isSpending': this_transaction['isSpending'],
			'isCashIn': this_transaction['isCashIn'], # to determine whether `amount` should be positive or negative
		}
		transactions_output.append(this_transaction_filtered)

	out = [summary, transactions_output]
	return out

def reshape_transactions(transactions):
	eventual_output = []
	for i in transactions:
		this_transaction_list = []
		for key in i.keys():
			this_transaction_list.append(i[key])
		eventual_output.append(this_transaction_list)
	return eventual_output

def main():
	# Check Google credentials
	creds = None
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'credentials.json', SCOPES)
			creds = flow.run_local_server(port=0)
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)
	service = build('sheets', 'v4', credentials=creds)


	# Download PC data
	pc_data = import_pc_data()
	summary_data = pc_data[0]
	transaction_data = pc_data[1]

	# for testing # pc_data = {'networth': 486483.22, 'investmentAccountsTotal': 394698.19}
	
	networth = summary_data['networth']
	investments = summary_data['investmentAccountsTotal']

	# reshape transaction data
	# returns a list of lists, where each sub-list is just the transaction values
	eventual_output = reshape_transactions(transaction_data)

	# read sheet to make sure we have data
	sheet = service.spreadsheets()
	range = SUMMARY_SHEET_NAME + '!A:C'
	result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
								range=range).execute()
	values = result.get('values', [])

	# if we don't yet have a row for this month, we need to insert one
	# get max row to see what month it is
	rows = result.get('values', [])
	max_row = len(rows)
	print(f'{max_row} rows retrieved.')

	# check max month to see if it's the current one
	max_date = values[max_row-1][0]
	max_month = max_date.split(' ')
	max_month = max_month[0]  # should be "September"
	print(f"here's the max date we have:  {max_date}")

	# get current month
	current_date = datetime.now()
	current_month = current_date.strftime("%B") # e.g. "August"
	current_year = str(current_date.strftime("%Y")) # e.g. "2020"

	is_current_month_already_present = current_month == max_month

	if is_current_month_already_present:
		# select the last row
		print("we already have a row for this month, so we'll just overwrite the values")
		summary_sheet_range =  SUMMARY_SHEET_NAME + '!A' + str(max_row) + ':C' + str(max_row)

	else:
		# insert a new row at the bottom for the current month
		print("we need to insert a new row for this month")

		summary_sheet_range = SUMMARY_SHEET_NAME + '!A' + str(max_row+1) + ':C' + str(max_row+1)


	# if we didn't get anything from PC API, then we don't upload anything
	if not values:
		print('No data found.')
	else:
		# upload summary data
		print("uploading summary data...")
		summary_body = {
			"values": [
				[
					current_month + ' ' + current_year,
					networth,
					investments
				]
			],
			"majorDimension": "ROWS"
		}
		result = service.spreadsheets().values().update(
			spreadsheetId=SPREADSHEET_ID, range=summary_sheet_range,
			valueInputOption='USER_ENTERED', body=summary_body).execute()
		print(result)


		# upload transactions data
		transactions_range = '!A2:I'
		transactions_sheet_range = TRANSACTIONS_SHEET_NAME + transactions_range

		print("uploading transactions...")
		transactions_body = {
			"values": eventual_output,
			"majorDimension": "ROWS"
		}
		result = service.spreadsheets().values().update(
			spreadsheetId=SPREADSHEET_ID, range=transactions_sheet_range,
			valueInputOption='USER_ENTERED', body=transactions_body).execute()
		print(result)

if __name__ == '__main__':
	main()
