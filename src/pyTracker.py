import email, ollama, json, os, base64, re, gspread, pandas as pd, difflib, logging
from email.header import decode_header
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/gmail.modify',
    'https://www.googleapis.com/auth/gmail.labels'
]


def get_credentials():
    config = _load_config()
    output_dir = config['output_dir']
    config_dir = "config"
    token_path = os.path.join(output_dir, 'token.json')
    
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_info(
            json.load(open(token_path)), SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                os.path.join(config_dir, 'credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open(token_path, 'w') as token:
            token.write(creds.to_json())
    
    return creds


def _load_config():
    with open('config/email_config.json', 'r') as f:
        return json.load(f)

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
            print(f"API returned {len(batch_messages)} messages in this batch")

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
                headers = {header['name']: header['value'] for header in msg['payload']['headers']}
                subject = headers.get('Subject', 'No Subject')
                raw_date = headers.get('Date', 'Unknown Date')
                sender = headers.get('From', 'Unknown Sender')
                # Convert date format to "Month Day, Year"
                try:
                    parsed_date = email.utils.parsedate_to_datetime(raw_date)
                    formatted_date = parsed_date.strftime("%m/%d/%Y")  
                except Exception:
                    formatted_date = raw_date  # Fallback if parsing fails

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
                
                body = re.sub(r'\s+', ' ', body).strip()
                # Add email to list with internalDate for sorting
                email_list.append({
                    'id': message['id'],
                    'subject': subject,
                    'date': formatted_date,
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

def getOllamaResponse(email,model):
    response = ollama.chat(model=model, messages=[
    {
        'role': 'system',
        'content': 'You are a JSON extraction tool ONLY. You must NEVER provide explanations, descriptions, or any text outside the requested JSON format. ONLY output valid JSON inside triple backticks.'
    },
    {
        'role': 'user',
        'content': '''
    CRITICAL INSTRUCTIONS:
    1. ONLY return a JSON object in the EXACT format shown below
    2. DO NOT include any explanations or descriptions
    3. DO NOT describe the HTML structure
    4. DO NOT engage in conversation
    5. Triple backticks MUST wrap your JSON response
    
    EXTRACT these fields:
    - "Job Name": Position title with ID if present
    - "Company": Company name (extract from domain or signature if needed)
    - "Status": EXACTLY one of: "Received", "Rejected", "Reviewing", or "Draft"
    
    STATUS DEFINITIONS:
    - "Received": Initial application acknowledgements, thank you messages
    - "Rejected": Clear rejections ("not moving forward", "other candidates", etc)
    - "Reviewing": Interview invites, being under consideration
    - "Draft": Only when status is completely unclear
    
    YOUR RESPONSE MUST BE ONLY:
    ```
    {
        "Job Name": "extracted job title",
        "Company": "extracted company name",
        "Status": "one of the allowed status values"
    }
    ```
    
    EMAIL:
    ''' + email,
    }])
    return response['message']['content']

def saveSheet(sh):
    values = sh.get_all_values()
    records = sh.get_all_records()
    df = pd.DataFrame(records)
    df.to_csv('config/data.csv', index=False)
    return df

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


def get_label_id(service, label_name):
    """Get the ID of a label, creating it if necessary."""
    labels = service.users().labels().list(userId='me').execute().get('labels', [])
    for label in labels:
        if label['name'].lower() == label_name.lower():
            return label['id']
    
    # If the label does not exist, create it
    label = {
        'name': label_name,
        'labelListVisibility': 'labelShow',
        'messageListVisibility': 'show'
    }
    created_label = service.users().labels().create(userId='me', body=label).execute()
    return created_label['id']

def add_label_to_emails(service, email_ids, label_id):
    """Add a label to a batch of emails."""
    body = {
        'ids': email_ids,
        'addLabelIds': [label_id],
        'removeLabelIds': []
    }
    service.users().messages().batchModify(userId='me', body=body).execute()

def updateSpreadsheet(worksheet, data):
    worksheet.batch_clear(['A3:H'])
    data = data.fillna('')
    data = data.map(lambda x: '' if pd.isna(x) else x)
    lists = data.values.tolist()
    worksheet.update(range_name='A2', values=lists)

def log_parse_failure(email_data, ai_response):
    """
    Log a parsing failure with clear separation between email and AI response.
    
    Args:
        email_data: The original email dictionary
        ai_response: The AI's response string
    """
    separator = "\n\n\n"  # 3 line breaks as separator
    
    email_formatted = json.dumps(email_data, indent=2)
    logging.info(f"======= FAILED EMAIL PARSING =======")
    logging.info(f"ORIGINAL EMAIL DATA:{separator}{email_formatted}{separator}")
    logging.info(f"AI RESPONSE:{separator}{ai_response}{separator}")
    logging.info(f"====================================")

def main():
    _CONFIG = _load_config()
    EMAIL = _CONFIG['email']
    APP_KEY = _CONFIG['appKey']
    SHEET_ID = _CONFIG['sheetID']
    model = _CONFIG['model_version']
    creds = get_credentials()
    logging.basicConfig(filename='output.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

    gmail_service = build('gmail', 'v1', credentials=creds)

    emails = get_emails_with_label(gmail_service, include_label='Internships', exclude_label='processed')
    if not emails:
        return None

    gc = gspread.authorize(creds)
    spreadsheet = gc.open_by_key(SHEET_ID)
    sh = spreadsheet.worksheet("Applications")

    df = saveSheet(sh)
    
    # Get the label ID for 'processed' to mark emails after processing
    processed_label_id = get_label_id(gmail_service, 'processed')
    emails_to_label = []  # To keep track of emails we've processed
    faulty_emails = []  # To keep track of faulty emails
    
    for email in emails:
        entry = {'Status':'','Company':'','Date Applied':'','Last Applied':'','Link':'','Role':'','Last Updated':''}
        ollamaResponse = getOllamaResponse(str(email),model)
        current = getJSON(ollamaResponse)
        
        if current is None:
            print(f"Failed to extract JSON from response for email with subject: {email.get('subject', 'Unknown subject')}")
            log_parse_failure(email, ollamaResponse)
            faulty_emails.append(email['id'])
            continue  # Skip to the next email
        
        # Only proceed if we have a valid response
        try:
            matches = difflib.get_close_matches(current['Company'], df['Company'], n=1, cutoff=0.8)
            if not matches:
                entry.update({
                    'Status': current['Status'],
                    'Company': current['Company'],
                    'Role': current['Job Name'],
                    'Last Updated': email['date'],
                    'Date Applied': email['date']
                })
                df.loc[len(df)] = entry
            else:
                match_index = df[df['Company'] == matches[0]].index[0]
                df.at[match_index, 'Status'] = current['Status']
                df.at[match_index, 'Role'] = current['Job Name']
                df.at[match_index, 'Last Updated'] = email['date']
            
            # Mark this email for labeling as processed
            emails_to_label.append(email['id'])
            
        except KeyError as e:
            print(f"Missing key in extracted data: {e} for email with subject: {email.get('subject', 'Unknown subject')}")
            print(f"Extracted data: {current}")
            faulty_emails.append(email['id'])
        except Exception as e:
            print(f"Unexpected error processing email: {e}")
            # May want to not mark as processed so it can be retried
    
    # Update the spreadsheet with the modified dataframe
    df.to_csv('config/data.csv', index=False)
    updateSpreadsheet(sh, df)
    
    # Mark all successfully processed emails
    if emails_to_label:
        add_label_to_emails(gmail_service, emails_to_label, processed_label_id)
        print(f"Marked {len(emails_to_label)} emails as processed")

if __name__ == "__main__":
    main()
