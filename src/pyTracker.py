import os, json
import ollama


# with open('config/email_config.json', 'r') as f:
#     config = json.load(f)

# SPREADSHEET_ID = config['sheetID']
# OUTPUT_DIR = config['output_dir']
# email = config['email']
# appKey = config['appKey']

# SCOPES = [
#     'https://www.googleapis.com/auth/spreadsheets',
#     'https://www.googleapis.com/auth/gmail.readonly',
#     'https://www.googleapis.com/auth/gmail.modify',
#     'https://www.googleapis.com/auth/gmail.labels'
# ]

email = """"""
response = ollama.chat(model='llama3:8b', messages=[
    {
        'role': 'user',
        'content': '''
You are a data entry assistant extracting job application information from emails.

TASK: Extract ONLY these three fields and return them as a clean JSON object:
1. "Job Name": The position title with ID if present (example: "J11032 Procurement Data Analyst Intern")
2. "Company": The company name (look for domain names in email addresses or URLs if not explicitly stated)
3. "Status": MUST be exactly one of: "Received", "Rejected", or "Reviewing"

STATUS RULES (very important):
- "Received": Use when the email acknowledges receipt of an application (phrases like "thank you for applying", "application received", etc.)
- "Reviewing": Use ONLY when the email explicitly states they are CURRENTLY reviewing the application
- "Rejected": Use ONLY when the email clearly states the application was not accepted

IMPORTANT: 
- Any email that says "thank you for applying" should be classified as "Received"
- If the company name appears in a URL (like company.com), extract it
- Always include the job ID if present (like J11032)

EMAIL:
''' + email,
    },
])
print(response['message']['content'])


