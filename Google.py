import pickle
import os
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class Gmail:
	def __init__(self, client_secret_file):
		CLIENT_SECRET_FILE = client_secret_file
		API_SERVICE_NAME = 'gmail'
		API_VERSION = 'v1'
		SCOPES = ['https://mail.google.com/']
		
		cred = None
		working_dir = os.getcwd()			# extract the working directory
		token_dir = 'token files'			# token folder name
		pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'

		# Check if token dir exists first, if not, create the folder
		if not os.path.exists(os.path.join(working_dir, token_dir)):
			os.mkdir(os.path.join(working_dir, token_dir))

		# Check if a token exist, load the credential data
		if os.path.exists(os.path.join(working_dir, token_dir, pickle_file)):
			with open(os.path.join(working_dir, token_dir, pickle_file), 'rb') as token:
				cred = pickle.load(token)

		# If the credential data doesn't exist or valid then try to request a new one
		if not cred or not cred.valid:
			if cred and cred.expired and cred.refresh_token:
				cred.refresh(Request())
			else:
				flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
				cred = flow.run_local_server()

			# Store the token
			with open(os.path.join(working_dir, token_dir, pickle_file), 'wb') as token:
				pickle.dump(cred, token)

		# Try to connect to the Gmail service with API
		try:
			self.service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
			emailAddress = self.service.users().getProfile(userId='me').execute()['emailAddress']
			print('Successfully connected to %s' %emailAddress)
		except Exception as e:
			print(e)
			print(f'Failed to create service instance for {API_SERVICE_NAME}')
			os.remove(os.path.join(working_dir, token_dir, pickle_file))
	
	def Send(self, receiver, subject, body):
		# Create a MIMEMultipart object and add the reciever's 
		# email address and the email's subject to it
		mimeMessage = MIMEMultipart()
		mimeMessage['to'] = receiver 
		mimeMessage['subject'] = subject
		
		# Convert the email's text to binary, decode, and send it
		mimeMessage.attach(MIMEText(body, 'plain'))
		raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
		message = self.service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
		print(message)

	def Read_emails(self):
		result = self.service.users().messages().list(userId='me').execute()
		messages = result.get('messages')
		List_of_Dicts = []				# Create an empty list
		
		for message in messages:
			# Get the message from its id
			txt = self.service.users().messages().get(userId='me', id=message['id']).execute()
			
			# Use try-except to avoid any Errors
			try:
				# Get value of 'payload' from dictionary 'txt'
				payload = txt['payload']
				headers = payload['headers']
				body = txt['snippet']

				# Look for Subject, Sender, and Receiver Email in the headers
				# all header['name']s become to lowercase before any comparison
				for header in headers:
					name = header['name'].lower()
					value = header['value']
					
					if name == 'subject':
						subject = value
					elif name == 'from':
						sender = value
					elif name == 'to':
						receiver = value

				# Convert to a dictionary
				dict = {
					'id': message['id'],
					'subject': subject,
					'sender': sender,
					'receiver': receiver,
					'body': body
				}

				# Add the dictionary to the list
				List_of_Dicts.append(dict)
			except:
				pass
		# Convert the list of dictionaries to a DataFrame
		self.mails = pd.DataFrame(List_of_Dicts)

	def Search(self, keyword):

		# define a function to make all email's text (subject, body, etc.)
		# characters to lowercase and eliminate signs to have pure words
		# in the list after split them
		def normalize(str):
			str = str.lower().strip()
			extras = [',', ':', '.', '!', '?', '\n', '\t', '@', '$', '\'s']
			for extra in extras:
				str = str.replace(extra, ' ')
			
			return str

		keyword = keyword.lower()		# Convert the keyword to lowercase letters
		self.Read_emails()				# Read all mails
		self.mails['body_list'] = self.mails['body'].apply(lambda x: normalize(x).split(' '))				# Normalize the body and separate them to a list of words
		self.mails['subject_list'] = self.mails['subject'].apply(lambda x: normalize(x).split(' '))			# Normalize the subject and separate them to a list of words
		self.mails['body_keyword_count'] = self.mails['body_list'].apply(lambda x: x.count(keyword))		# Count the keyword on the body's text
		self.mails['subject_keyword_count'] = self.mails['subject_list'].apply(lambda x: x.count(keyword))	# Count the keyword on the subject's text
		self.mails['keyword_search_points'] = self.mails['body_keyword_count'] + 5*self.mails['subject_keyword_count']		# Let's assume subject is 5 times more important than the body
		search_result = self.mails.loc[self.mails['keyword_search_points']>0].sort_values(by='keyword_search_points', ascending=False)	# Filter and sort the results and points
		
		# Iterate on the results' DataFrame and display them separately
		for index, row in search_result.iterrows():
			print('------------------')
			print('id:', row['id'])
			print('sender:', row['sender'])
			print('receiver:', row['receiver'])
			print('subject:', row['subject'])
			print('body:', row['body'])
		
		if len(search_result) == 0:
			print('There is no email with the word %s' %keyword)
			
