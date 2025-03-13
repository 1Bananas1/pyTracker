import os, json, re, gspread, base64, ast
import ollama
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build



with open('config/email_config.json', 'r') as f:
    config = json.load(f)

SPREADSHEET_ID = config['sheetID']
OUTPUT_DIR = config['output_dir']
email = config['email']
appKey = config['appKey']
CONFIG_DIR = "config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "email_config.json")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels'
]

def get_emails_with_label(service, include_label='Internships', exclude_label='y'):
    """
    Retrieve emails with a specific label but excluding another label.
    Returns a list of dictionaries containing email details.
    
    Args:
        service: The Gmail API service instance
        include_label: The label emails must have (default 'Internships')
        exclude_label: The label emails must not have (default 'y')
        
    Returns:
        List of dictionaries with email details (subject, body, date, etc.)
    """
    try:
        # Construct query to get emails with label x but not label y
        query = f"label:{include_label} -label:{exclude_label}"
        print(f"Searching for emails with query: {query}")
        
        messages = []
        next_page_token = None
        
        # Loop to get all messages using pagination
        while True:
            # Get list of messages matching the query
            results = service.users().messages().list(
                userId='me', 
                q=query,
                pageToken=next_page_token,
                maxResults=500  # Request maximum allowed per page
            ).execute()
            
            batch_messages = results.get('messages', [])
            if batch_messages:
                messages.extend(batch_messages)
                print(f"Retrieved {len(messages)} emails so far...")
            
            # Check if there are more pages
            next_page_token = results.get('nextPageToken')
            if not next_page_token:
                break
        
        if not messages:
            print('No emails found with the specified labels.')
            return []
        
        print(f"Found a total of {len(messages)} emails matching criteria.")
        
        # Process each email to extract details
        email_list = []
        for i, message in enumerate(messages):
            try:
                if i % 20 == 0:  # Progress update every 20 emails
                    print(f"Processing email {i+1}/{len(messages)}...")
                
                msg = service.users().messages().get(userId='me', id=message['id']).execute()
                
                # Extract headers
                headers = {}
                for header in msg['payload']['headers']:
                    headers[header['name']] = header['value']
                
                # Get subject, date, and sender
                subject = headers.get('Subject', 'No Subject')
                date = headers.get('Date', 'Unknown Date')
                sender = headers.get('From', 'Unknown Sender')
                
                # Extract body content
                body = ""
                
                if 'parts' in msg['payload']:
                    for part in msg['payload']['parts']:
                        if part['mimeType'] == 'text/plain':
                            if 'data' in part['body']:
                                body_data = part['body']['data']
                                body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                                break
                elif 'body' in msg['payload'] and 'data' in msg['payload']['body']:
                    body_data = msg['payload']['body']['data']
                    body = base64.urlsafe_b64decode(body_data).decode('utf-8')
                
                # Add email to list with internalDate for sorting
                email_list.append({
                    'id': message['id'],
                    'subject': subject,
                    'date': date,
                    'internal_date': int(msg.get('internalDate', 0)),
                    'sender': sender,
                    'body': body
                })
            except Exception as e:
                print(f"Error processing message {message['id']}: {e}")
        
        # Sort the list by internal_date (oldest first)
        if email_list:
            email_list.sort(key=lambda x: x['internal_date'])
        
        return email_list
    
    except Exception as e:
        print(f"Error retrieving emails: {e}")
        if "insufficient authentication scopes" in str(e):
            print("\nPermission error: Your authentication token doesn't have the necessary Gmail permissions.")
            print("Please update your SCOPES list to include Gmail permissions and regenerate your token.")
            print("Required scopes: https://www.googleapis.com/auth/gmail.readonly or https://www.googleapis.com/auth/gmail.modify")
        return []

def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens
    token_path = os.path.join(OUTPUT_DIR, 'token.json')
    
    # Check if token.json exists with valid credentials
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_info(
            json.load(open(token_path)), SCOPES)
    
    # If there are no valid credentials, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(CONFIG_DIR, 'credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    return creds



def getOllamaResponse(email):
    response = ollama.chat(model='llama3:8b', messages=[
    {
        'role': 'user',
        'content': '''
    You are a data entry assistant extracting job application information from emails.

    TASK: Extract ONLY these three fields:
    1. "Job Name": The position title with ID if present
    2. "Company": The company name 
    3. "Status": MUST be exactly one of: "Received", "Rejected", "Reviewing", or "Draft"

    Provide your answer in this exact format with NO additional explanation:
    ```
    {
        "Job Name": "extracted job title",
        "Company": "extracted company name",
        "Status": "one of the allowed status values"
    }
    ```

    EMAIL:
    ''' + email,
        },
    ])
    response_message = response['message']['content']
    return response_message

def getJSON(response_message):
    # First, try to parse the entire response as JSON directly
    try:
        import json
        data = json.loads(response_message.strip())
        return data
    except json.JSONDecodeError:
        pass
    
    # If that fails, look for JSON in code blocks
    match = re.search(r"```(?:json)?(.*?)```", response_message, re.DOTALL)
    if match:
        json_str = match.group(1).strip()
        try:
            import json
            data = json.loads(json_str)
            return data
        except json.JSONDecodeError:
            pass
    
    # If both methods fail, try to find a JSON-like structure
    match = re.search(r"\{[\s\S]*?\}", response_message, re.DOTALL)
    if match:
        json_str = match.group(0)
        try:
            import json
            data = json.loads(json_str)
            return data
        except json.JSONDecodeError:
            pass
    
    # Last resort: regex extraction for key-value pairs
    extracted_data = {}
    
    # Modified regex patterns to handle empty values
    job_name_match = re.search(r'"Job Name"\s*:\s*"([^"]*)"', response_message)
    company_match = re.search(r'"Company"\s*:\s*"([^"]*)"', response_message)
    status_match = re.search(r'"Status"\s*:\s*"([^"]*)"', response_message)
    
    if job_name_match:
        extracted_data["Job Name"] = job_name_match.group(1)
    if company_match:
        extracted_data["Company"] = company_match.group(1)
    if status_match:
        extracted_data["Status"] = status_match.group(1)
    
    if extracted_data and "Company" in extracted_data and "Status" in extracted_data:
        # Ensure Job Name exists even if empty
        if "Job Name" not in extracted_data:
            extracted_data["Job Name"] = ""
        return extracted_data
    
    print("Could not extract data from response:")
    print(response_message)
    return None

def normalize_company_name(company_name):
    """
    Normalizes company names to handle inconsistencies in capitalization and formatting.
    Example: "WEX Inc." and "Wex Inc." become "wex"
    """
    if not company_name or company_name.lower() == "unknown":
        return "unknown"
    
    # Remove common suffixes and normalize case
    normalized = company_name.lower()
    suffixes = [" inc", " inc.", " llc", " llc.", " corp", " corp.", 
                " corporation", " co", " co.", " company", " ltd", " ltd."]
    
    for suffix in suffixes:
        if normalized.endswith(suffix):
            normalized = normalized[:-len(suffix)]
            break
    
    # Remove extra spaces and punctuation (including hyphens)
    normalized = re.sub(r'[^\w\s]', '', normalized).strip()
    return normalized

def check_company_in_backend(gc, spreadsheet_id, company_name):
    """
    Check if the given company name exists in the 'Backend' sheet,
    using normalization to handle inconsistencies.
    
    Returns tuple: (exists, exact_match_name)
    - exists: True if a normalized match exists
    - exact_match_name: The exact company name from the sheet that matched
    """
    try:
        # Normalize the input company name
        normalized_company = normalize_company_name(company_name)
        
        # Open the spreadsheet
        spreadsheet = gc.open_by_key(spreadsheet_id)
        
        # Select the 'Backend' sheet
        backend_sheet = spreadsheet.worksheet('Backend')
        
        # Get all values in column A starting from row 2
        company_list = backend_sheet.col_values(1)[1:]  # Skip the header
        
        # Check for normalized matches
        for existing_company in company_list:
            if normalize_company_name(existing_company) == normalized_company:
                return True, existing_company
        
        return False, None
    
    except Exception as e:
        print(f"Error checking company in 'Backend' sheet: {e}")
        return False, None

def main():
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)
    gc = gspread.authorize(creds)
    gmail_service = build('gmail', 'v1', credentials=creds)
    sheet_name = 'Applications'
    back_end = 'Backend'
    
    email_list = get_emails_with_label(gmail_service, 'Internships', 'ProcessedI')
    for i, email in enumerate(email_list, 1):
        print(f"\nProcessing email {i}/{len(email_list)}: {email['subject'][:50]}...")
        
        current = getJSON(getOllamaResponse(str(email)))
        if current is None:
            print("Failed to parse JSON from email response.")
            continue
            
        # The updated check_company_in_backend returns a tuple, but you're using it as a boolean
        company_exists, existing_name = check_company_in_backend(gc, SPREADSHEET_ID, current['Company'])
        
        print(f"Extracted company: '{current['Company']}'")
        
        if company_exists:
            print(f"✓ Matched with existing company: '{existing_name}'")
            current['Company'] = existing_name  # Use the existing name format for consistency
        else:
            print(f"+ Adding new company: '{current['Company']}'")
            try:
                spreadsheet = gc.open_by_key(SPREADSHEET_ID)
                backend_sheet = spreadsheet.worksheet('Backend')
                backend_sheet.append_row([current['Company']])
                print("Successfully added to backend sheet")
            except Exception as e:
                print(f"Error adding company to backend: {e}")

if __name__ == '__main__':
    main()