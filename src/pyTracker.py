import os, json



with open('config/email_config.json', 'r') as f:
    config = json.load(f)

SPREADSHEET_ID = config['sheetID']
OUTPUT_DIR = config['output_dir']
email = config['email']
appKey = config['appKey']

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels'
]