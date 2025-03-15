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

def get_sheet_data_cached(gc, spreadsheet_id, cache_duration=600):
    """
    Get all required data from Google Sheets in a single batch at the beginning of a run.
    Uses caching to minimize API calls.
    
    Args:
        gc: Authorized gspread client
        spreadsheet_id: ID of the Google Sheet
        cache_duration: Duration in seconds for which the cache remains valid (default: 10 minutes)
    
    Returns:
        Dictionary containing sheets, worksheet data, and metadata
    """
    import os
    import pickle
    import time
    from datetime import datetime
    
    # Define the cache file path
    cache_file = os.path.join("config", "sheets_cache.pkl")
    cache_valid = False
    
    # Check if we have a valid cache
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'rb') as f:
                cache_data = pickle.load(f)
                
            # Check if cache is still valid
            if cache_data.get('timestamp', 0) > time.time() - cache_duration:
                print(f"Using cached spreadsheet data (valid for {int((cache_data['timestamp'] + cache_duration - time.time()) / 60)} more minutes)")
                cache_valid = True
                return cache_data['data']
            else:
                print("Cache expired, fetching fresh data...")
        except Exception as e:
            print(f"Error reading cache: {e}")
    
    # If no valid cache, fetch all data in a single batch
    print("Fetching all spreadsheet data...")
    
    try:
        # Open the spreadsheet just once
        spreadsheet = gc.open_by_key(spreadsheet_id)
        
        # Get all required worksheets
        applications_sheet = spreadsheet.worksheet('Applications')
        backend_sheet = spreadsheet.worksheet('Backend')
        
        # Fetch all data at once to minimize API calls
        applications_data = applications_sheet.get_all_values()
        backend_data = backend_sheet.get_all_values()
        
        # Prepare data structure to return and cache
        sheet_data = {
            'spreadsheet': spreadsheet,  # Keep reference to avoid reopening
            'worksheets': {
                'Applications': applications_sheet,
                'Backend': backend_sheet
            },
            'data': {
                'Applications': applications_data,
                'Backend': backend_data
            },
            'metadata': {
                'last_updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'row_counts': {
                    'Applications': len(applications_data),
                    'Backend': len(backend_data)
                }
            }
        }
        
        # Cache the data for future use
        try:
            os.makedirs(os.path.dirname(cache_file), exist_ok=True)
            with open(cache_file, 'wb') as f:
                pickle.dump({
                    'timestamp': time.time(),
                    'data': sheet_data
                }, f)
            print("Spreadsheet data cached successfully")
        except Exception as e:
            print(f"Error caching spreadsheet data: {e}")
        
        return sheet_data
    except Exception as e:
        print(f"Error fetching spreadsheet data: {e}")
        
        # Fall back to cache if available, even if expired
        if os.path.exists(cache_file):
            try:
                with open(cache_file, 'rb') as f:
                    cache_data = pickle.load(f)
                print("Falling back to expired cache due to error")
                return cache_data['data']
            except:
                pass
                
        # Return empty structure if all else fails
        return {
            'spreadsheet': None,
            'worksheets': {},
            'data': {'Applications': [], 'Backend': []},
            'metadata': {'last_updated': 'never', 'row_counts': {'Applications': 0, 'Backend': 0}}
        }

def process_applications_from_cache(sheet_data):
    """
    Process application data from cache instead of making repeated API calls.
    Returns a dictionary mapping company_id/job_id tuples to application details.
    """
    application_map = {}
    
    # Skip if data is missing
    if not sheet_data or 'data' not in sheet_data or 'Applications' not in sheet_data['data']:
        return application_map
    
    applications_data = sheet_data['data']['Applications']
    
    # Skip if only header or empty
    if len(applications_data) <= 1:
        return application_map
    
    # First row is headers
    headers = applications_data[0]
    
    # Find index positions for required columns
    company_idx = headers.index('Company') if 'Company' in headers else -1
    role_idx = headers.index('Role') if 'Role' in headers else -1
    company_id_idx = headers.index('Company ID') if 'Company ID' in headers else -1
    job_id_idx = headers.index('Job ID') if 'Job ID' in headers else -1
    
    # Process each application row
    for i, row in enumerate(applications_data[1:], 1):  # Skip header, use 1-based index
        # Skip rows that don't have all required fields
        if len(row) <= max(company_idx, role_idx, company_id_idx, job_id_idx):
            continue
            
        # Get values for keys
        company_id = row[company_id_idx] if company_id_idx >= 0 and company_id_idx < len(row) else ""
        job_id = row[job_id_idx] if job_id_idx >= 0 and job_id_idx < len(row) else ""
        company = row[company_idx] if company_idx >= 0 and company_idx < len(row) else ""
        role = row[role_idx] if role_idx >= 0 and role_idx < len(row) else ""
        
        # Create both ID-based and name-based keys for matching
        id_key = (company_id, job_id) if company_id and job_id else None
        name_key = (normalize_company_name(company), role.lower().strip()) if company and role else None
        
        # Store application data with both keys if available
        application_data = {
            'row_index': i + 1,  # +1 for header row
            'data': row
        }
        
        if id_key:
            application_map[id_key] = application_data
        if name_key:
            application_map[name_key] = application_data
    
    print(f"Processed {len(application_map)} application entries from cache")
    return application_map

def process_companies_from_cache(sheet_data):
    """
    Process company data from cache instead of making repeated API calls.
    Returns a dictionary mapping normalized company names to their details.
    """
    company_map = {}
    
    # Skip if data is missing
    if not sheet_data or 'data' not in sheet_data or 'Backend' not in sheet_data['data']:
        return company_map
    
    backend_data = sheet_data['data']['Backend']
    
    # Skip header row
    if len(backend_data) <= 1:
        return company_map
    
    # Process company data
    for row in backend_data[1:]:  # Skip header row
        if len(row) < 2:
            continue
            
        company_name = row[0]
        company_id = row[1] if len(row) > 1 and row[1] else ""
        
        # Normalize the company name for case-insensitive matching
        normalized_name = normalize_company_name(company_name)
        
        # Store both original name and ID
        company_map[normalized_name] = {
            'original_name': company_name,
            'company_id': company_id
        }
    
    print(f"Processed {len(company_map)} companies from cache")
    return company_map

def batch_update_sheets(gc, sheet_data, updates):
    """
    Perform all Google Sheet updates in a single batch to minimize API calls.
    
    Args:
        gc: Authorized gspread client
        sheet_data: Cached sheet data from get_sheet_data_cached
        updates: Dictionary with update operations
    
    Returns:
        Results of the batch operations
    """
    # Skip if no updates
    if not updates:
        print("No updates to process")
        return None
    
    results = {
        'new_companies': 0,
        'updated_applications': 0,
        'new_applications': 0,
        'errors': []
    }
    
    try:
        # Get spreadsheet and worksheet references
        spreadsheet = sheet_data.get('spreadsheet')
        if not spreadsheet:
            # If we don't have a cached reference, get a new one
            spreadsheet = gc.open_by_key(updates.get('spreadsheet_id', ''))
        
        # Get the worksheets we need
        backend_sheet = sheet_data.get('worksheets', {}).get('Backend')
        applications_sheet = sheet_data.get('worksheets', {}).get('Applications')
        
        if not backend_sheet:
            backend_sheet = spreadsheet.worksheet('Backend')
        if not applications_sheet:
            applications_sheet = spreadsheet.worksheet('Applications')
        
        # 1. Add new companies in a single batch
        if 'new_companies' in updates and updates['new_companies']:
            try:
                backend_sheet.append_rows(updates['new_companies'])
                results['new_companies'] = len(updates['new_companies'])
                print(f"Added {len(updates['new_companies'])} new companies in a single batch")
            except Exception as e:
                results['errors'].append(f"Error adding companies: {str(e)}")
                print(f"Error adding companies: {e}")
        
        # 2. Update existing applications in a single batch
        if 'application_updates' in updates and updates['application_updates']:
            try:
                spreadsheet_service = build('sheets', 'v4', credentials=gc.auth)
                batch_update_request = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': updates['application_updates']
                }
                spreadsheet_service.spreadsheets().values().batchUpdate(
                    spreadsheetId=updates.get('spreadsheet_id', ''),
                    body=batch_update_request
                ).execute()
                results['updated_applications'] = len(updates['application_updates']) // 2  # Divide by 2 since each update requires 2 cells
                print(f"Updated {results['updated_applications']} applications in a single batch")
            except Exception as e:
                results['errors'].append(f"Error updating applications: {str(e)}")
                print(f"Error updating applications: {e}")
        
        # 3. Add new applications in a single batch
        if 'new_applications' in updates and updates['new_applications']:
            try:
                applications_sheet.append_rows(updates['new_applications'])
                results['new_applications'] = len(updates['new_applications'])
                print(f"Added {len(updates['new_applications'])} new applications in a single batch")
            except Exception as e:
                results['errors'].append(f"Error adding new applications: {str(e)}")
                print(f"Error adding new applications: {e}")
        
        # 4. Save any metadata updates if needed
        if 'metadata_updates' in updates and updates['metadata_updates']:
            try:
                # Implementation depends on how you want to store metadata
                pass
            except Exception as e:
                results['errors'].append(f"Error updating metadata: {str(e)}")
        
        return results
    
    except Exception as e:
        results['errors'].append(f"General batch update error: {str(e)}")
        print(f"Error in batch update: {e}")
        return results

def normalize_company_name(company_name):
    """
    Normalizes company names to handle inconsistencies in capitalization and formatting.
    Example: "WEX Inc." and "Wex Inc." become "wex"
    """
    import re
    
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

# Example of Using the Optimized Google Sheets Functions

def process_emails_optimized(gc, gmail_service, spreadsheet_id):
    """
    Process emails with optimized Google Sheets API usage.
    This example shows how to dramatically reduce the number of API calls.
    """
    # Step 1: Get all emails that need processing
    email_list = get_emails_with_label_batch(gmail_service, 'Internships', 'y')
    
    if not email_list:
        print("No new emails to process.")
        return
    
    print(f"Found {len(email_list)} new emails to process.")
    
    # Step 2: Get all spreadsheet data in ONE API call sequence instead of for EACH email
    sheet_data = get_sheet_data_cached(gc, spreadsheet_id)
    
    # Step 3: Process company and application data from cache - NO API calls
    company_map = process_companies_from_cache(sheet_data)
    application_map = process_applications_from_cache(sheet_data)
    
    # Step 4: Process all emails using cached data
    # Prepare containers for all updates so we can do them in a single batch
    updates = {
        'spreadsheet_id': spreadsheet_id,
        'new_companies': [],
        'application_updates': [],
        'new_applications': []
    }
    
    # Track unique entries to avoid duplicates
    processed_companies = set()
    processed_applications = set()
    
    # Process each email
    for email in email_list:
        try:
            print(f"Processing email: {email['subject'][:50]}...")
            
            # Extract data from email
            parsed_data = getJSON(getOllamaResponse(email))
            
            if parsed_data is None:
                print("Failed to parse data from email.")
                continue
            
            # Extract company info
            company_name = parsed_data.get('Company', '')
            normalized_company = normalize_company_name(company_name)
            
            # Extract job info
            job_name = parsed_data.get('Job Name', '')
            status = parsed_data.get('Status', 'Received')
            
            # Check if company exists in our cached data - NO API call
            company_exists = normalized_company in company_map
            
            # Company processing
            if company_exists:
                # Use existing company data
                existing_data = company_map[normalized_company]
                company_name = existing_data['original_name']
                company_id = existing_data['company_id']
                print(f"✓ Matched with existing company: '{company_name}'")
            else:
                # Check if we've already processed this company in this batch
                if normalized_company in processed_companies:
                    # Find the company in our updates list
                    for i, company_entry in enumerate(updates['new_companies']):
                        if normalize_company_name(company_entry[0]) == normalized_company:
                            company_name = company_entry[0]
                            company_id = company_entry[1]
                            print(f"✓ Using already processed company: '{company_name}'")
                            break
                else:
                    # Add new company to batch updates
                    print(f"+ Adding new company: '{company_name}'")
                    company_id = f"COMP{len(company_map) + len(updates['new_companies']) + 1}"
                    
                    # Add to updates batch - will be sent in ONE API call later
                    updates['new_companies'].append([company_name, company_id])
                    processed_companies.add(normalized_company)
                    
                    # Update our local map for future reference
                    company_map[normalized_company] = {
                        'original_name': company_name,
                        'company_id': company_id
                    }
            
            # Application processing
            job_id = f"JOB{len(application_map) + len(updates['new_applications']) + 1}"
            
            # Create keys for checking existing applications
            application_key = (company_id, job_id)
            alternative_key = (normalized_company, job_name.lower().strip())
            
            # Check if this application already exists - NO API call
            existing_application = application_map.get(application_key)
            if not existing_application:
                existing_application = application_map.get(alternative_key)
            
            # Format dates
            current_date = datetime.now().strftime("%Y-%m-%d")
            
            # Application identity for deduplication
            application_identity = f"{company_id}:{job_name.lower().strip()}"
            
            if existing_application:
                # Update existing application if not already processed
                if application_identity in processed_applications:
                    print(f"✓ Already processed this application in current batch")
                    continue
                
                processed_applications.add(application_identity)
                
                row_index = existing_application['row_index']
                old_status = existing_application['data'][0] if len(existing_application['data']) > 0 else "Unknown"
                
                # Only update if status has changed
                if old_status != status:
                    # Add to batch updates - will be sent in ONE API call
                    updates['application_updates'].append({
                        'range': f'Applications!A{row_index}',
                        'values': [[status]]
                    })
                    updates['application_updates'].append({
                        'range': f'Applications!D{row_index}',
                        'values': [[current_date]]
                    })
                    
                    print(f"Updated existing application for '{company_name}' - '{job_name}'")
                else:
                    print(f"Status unchanged for '{company_name}' - '{job_name}', skipping update")
            else:
                # Add new application if not already processed
                if application_identity in processed_applications:
                    print(f"✓ Already added this application in current batch")
                    continue
                
                processed_applications.add(application_identity)
                
                # Create new application record
                new_application = [
                    status,                # Status
                    company_name,          # Company
                    current_date,          # Date Applied
                    current_date,          # Last Updated
                    "",                    # Link
                    job_name,              # Role
                    company_id,            # Company ID
                    job_id                 # Job ID
                ]
                
                # Add to batch updates - will be sent in ONE API call
                updates['new_applications'].append(new_application)
                
                print(f"Added new application for '{company_name}' - '{job_name}'")
                
        except Exception as e:
            print(f"Error processing email: {e}")
    
    # Step 5: Send ALL updates in a single batch - dramatically reduces API calls
    if (updates['new_companies'] or updates['application_updates'] or updates['new_applications']):
        print("\nSending all updates in a single batch...")
        results = batch_update_sheets(gc, sheet_data, updates)
        
        # Print results
        print("\n=== Update Results ===")
        print(f"New companies added: {results['new_companies']}")
        print(f"Existing applications updated: {results['updated_applications']}")
        print(f"New applications added: {results['new_applications']}")
        
        if results['errors']:
            print("\nErrors encountered:")
            for error in results['errors']:
                print(f"- {error}")
    else:
        print("No updates needed")
    
    print("\nProcessing complete!")

def get_or_create_processed_label(gmail_service):
    """
    Get the ID of the 'Processed' label, or create it if it doesn't exist.
    This avoids repeatedly checking if the label exists.
    """
    try:
        # Get all labels
        results = gmail_service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        
        # Look for an existing 'Processed' label
        for label in labels:
            if label['name'] == 'Processed':
                return label['id']
        
        # If not found, create it
        label_object = {
            'name': 'Processed',
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show'
        }
        created_label = gmail_service.users().labels().create(
            userId='me', 
            body=label_object
        ).execute()
        
        return created_label['id']
    
    except Exception as e:
        print(f"Error managing 'Processed' label: {e}")
        return None

def get_unprocessed_emails(gmail_service, include_label='Internships', processed_label='Processed'):
    """
    Get emails with a specific label that haven't been processed yet.
    This is more reliable than date-based filtering.
    """
    try:
        # Get the processed label ID
        processed_label_id = get_or_create_processed_label(gmail_service)
        
        # Query for emails with include_label but not processed_label
        query = f"label:{include_label} -label:{processed_label}"
        print(f"Looking for unprocessed emails with query: {query}")
        
        # Fetch message IDs
        results = gmail_service.users().messages().list(
            userId='me', 
            q=query,
            maxResults=500
        ).execute()
        
        messages = results.get('messages', [])
        if not messages:
            print('No unprocessed emails found.')
            return []
        
        print(f"Found {len(messages)} unprocessed emails.")
        
        # Process emails in batches for efficiency
        email_list = []
        batch_size = 25
        
        for i in range(0, len(messages), batch_size):
            batch = messages[i:i+batch_size]
            print(f"Fetching details for batch {i//batch_size + 1}/{len(messages)//batch_size + 1}...")
            
            # Get email details for the batch
            for message in batch:
                try:
                    msg = gmail_service.users().messages().get(
                        userId='me', 
                        id=message['id']
                    ).execute()
                    
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
                    
                    # Add email to list
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
        
        # Sort the emails by date (oldest first)
        email_list.sort(key=lambda x: x['internal_date'])
        
        return email_list
    
    except Exception as e:
        print(f"Error retrieving emails: {e}")
        return []

def mark_emails_as_processed(gmail_service, email_ids):
    """
    Mark multiple emails as processed in a single batch operation.
    Much more efficient than marking one at a time.
    """
    if not email_ids:
        return
    
    try:
        # Get the processed label ID
        processed_label_id = get_or_create_processed_label(gmail_service)
        if not processed_label_id:
            print("Could not get or create 'Processed' label, emails won't be marked")
            return
        
        # Apply the label in a single batch operation
        batch_modify_request = {
            'ids': email_ids,
            'addLabelIds': [processed_label_id]
        }
        
        gmail_service.users().messages().batchModify(
            userId='me', 
            body=batch_modify_request
        ).execute()
        
        print(f"Marked {len(email_ids)} emails as processed")
    
    except Exception as e:
        print(f"Error marking emails as processed: {e}")




def main():
    """Main function with optimized Google Sheets API usage and email labeling"""
    # Get API credentials
    creds = get_credentials()
    gc = gspread.authorize(creds)
    gmail_service = build('gmail', 'v1', credentials=creds)
    
    # Get all emails that need processing - using label filtering instead of date
    email_list = get_unprocessed_emails(gmail_service, include_label='Internships', processed_label='Processed')
    
    if not email_list:
        print("No unprocessed emails to process.")
        return
    
    print(f"Found {len(email_list)} unprocessed emails to process.")
    
    # Track email IDs to mark as processed at the end
    processed_email_ids = []
    
    # Get all spreadsheet data in a single batch at the beginning
    try:
        print("Loading spreadsheet data...")
        spreadsheet = gc.open_by_key(SPREADSHEET_ID)
        backend_sheet = spreadsheet.worksheet('Backend')
        applications_sheet = spreadsheet.worksheet('Applications')
        
        # Get all company data once
        company_rows = backend_sheet.get_all_values()
        
        # Get all application data once (if needed)
        application_rows = applications_sheet.get_all_values()
        
        # Create a company map for quick lookups without API calls
        company_map = {}
        if len(company_rows) > 1:  # Skip header row
            for row in company_rows[1:]:
                if not row or not row[0]:
                    continue
                company_name = row[0]
                company_id = row[1] if len(row) > 1 and row[1] else ""
                normalized_name = normalize_company_name(company_name)
                company_map[normalized_name] = {
                    'original_name': company_name,
                    'company_id': company_id
                }
        
        # Create application map for quick lookups (if needed)
        application_map = {}
        if len(application_rows) > 1:  # Skip header row
            headers = application_rows[0]
            
            # Find index positions for required columns
            company_idx = headers.index('Company') if 'Company' in headers else -1
            role_idx = headers.index('Role') if 'Role' in headers else -1
            company_id_idx = headers.index('Company ID') if 'Company ID' in headers else -1
            job_id_idx = headers.index('Job ID') if 'Job ID' in headers else -1
            
            for i, row in enumerate(application_rows[1:], 2):  # +2 for 1-based indexing and header row
                if len(row) <= max(company_idx, role_idx):
                    continue
                
                company = row[company_idx] if company_idx >= 0 and company_idx < len(row) else ""
                role = row[role_idx] if role_idx >= 0 and role_idx < len(row) else ""
                
                if company and role:
                    key = (normalize_company_name(company), role.lower().strip())
                    application_map[key] = {
                        'row_index': i,
                        'data': row
                    }
        
        print(f"Loaded {len(company_map)} companies and {len(application_map)} applications.")
    except Exception as e:
        print(f"Error loading spreadsheet data: {e}")
        return
    
    # Prepare batch updates
    new_companies = []
    new_applications = []
    application_updates = []
    
    # Process all emails
    for i, email in enumerate(email_list, 1):
        try:
            print(f"\nProcessing email {i}/{len(email_list)}: {email['subject'][:50]}...")
            
            # Add this email to the list of processed emails
            processed_email_ids.append(email['id'])
            
            current = getJSON(getOllamaResponse(str(email)))
            if current is None:
                print("Failed to parse JSON from email response.")
                continue
                
            # Normalize company name for matching
            company_name = current.get('Company', '')
            normalized_company = normalize_company_name(company_name)
            
            # Extract job information
            job_name = current.get('Job Name', '')
            status = current.get('Status', 'Received')
            
            print(f"Extracted company: '{company_name}', job: '{job_name}', status: '{status}'")
            
            # Check if company exists in our cached data (NO API CALL)
            company_exists = normalized_company in company_map
            
            if company_exists:
                existing_data = company_map[normalized_company]
                existing_name = existing_data['original_name']
                company_id = existing_data.get('company_id', '')
                print(f"✓ Matched with existing company: '{existing_name}'")
                company_name = existing_name  # Use the existing name format for consistency
            else:
                print(f"+ Adding new company: '{company_name}'")
                # Add to batch instead of immediate API call
                company_id = f"COMP{len(company_map) + len(new_companies) + 1}"
                new_companies.append([company_name, company_id])
                
                # Update local map for future emails in this batch
                company_map[normalized_company] = {
                    'original_name': company_name,
                    'company_id': company_id
                }
            
            # Check if we have an existing application for this company and role
            application_key = (normalized_company, job_name.lower().strip())
            existing_application = application_map.get(application_key)
            
            # Format dates
            current_date = datetime.now().strftime("%Y-%m-%d")
            
            if existing_application:
                # Update existing application
                row_index = existing_application['row_index']
                old_status = existing_application['data'][0] if len(existing_application['data']) > 0 else ""
                
                # Only update if status has changed
                if status != old_status:
                    application_updates.append({
                        'range': f'Applications!A{row_index}',
                        'values': [[status]]
                    })
                    application_updates.append({
                        'range': f'Applications!D{row_index}',
                        'values': [[current_date]]
                    })
                    print(f"Updated application status from '{old_status}' to '{status}'")
                else:
                    print(f"Status unchanged for application, skipping update")
            else:
                # Create new application
                job_id = f"JOB{len(application_map) + len(new_applications) + 1}"
                
                # Create new application record
                new_application = [
                    status,          # Status
                    company_name,    # Company
                    current_date,    # Date Applied
                    current_date,    # Last Updated
                    "",              # Link
                    job_name,        # Role
                    company_id,      # Company ID
                    job_id           # Job ID
                ]
                
                new_applications.append(new_application)
                
                # Update our local map
                application_map[application_key] = {
                    'row_index': len(application_rows) + len(new_applications),
                    'data': new_application
                }
                
                print(f"Added new application for '{company_name}' - '{job_name}'")
        except Exception as e:
            print(f"Error processing email: {e}")
    
    # Perform batch updates at the end
    sheets_updated = False
    
    # Update companies
    if new_companies:
        try:
            print(f"\nAdding {len(new_companies)} new companies in a single batch...")
            backend_sheet.append_rows(new_companies)
            print(f"Successfully added {len(new_companies)} companies to backend sheet")
            sheets_updated = True
        except Exception as e:
            print(f"Error adding companies to backend: {e}")
    
    # Update existing applications
    if application_updates:
        try:
            print(f"\nUpdating {len(application_updates) // 2} applications in a single batch...")
            sheets_service = build('sheets', 'v4', credentials=creds)
            batch_update_request = {
                'valueInputOption': 'USER_ENTERED',
                'data': application_updates
            }
            sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body=batch_update_request
            ).execute()
            print(f"Successfully updated {len(application_updates) // 2} applications")
            sheets_updated = True
        except Exception as e:
            print(f"Error updating applications: {e}")
    
    # Add new applications
    if new_applications:
        try:
            print(f"\nAdding {len(new_applications)} new applications in a single batch...")
            applications_sheet.append_rows(new_applications)
            print(f"Successfully added {len(new_applications)} new applications")
            sheets_updated = True
        except Exception as e:
            print(f"Error adding new applications: {e}")
    
    # Mark all processed emails with the Processed label
    if processed_email_ids:
        try:
            print(f"\nMarking {len(processed_email_ids)} emails as processed...")
            mark_emails_as_processed(gmail_service, processed_email_ids)
        except Exception as e:
            print(f"Error marking emails as processed: {e}")
    
    if sheets_updated:
        print("\nAll spreadsheet updates completed successfully.")
    else:
        print("\nNo spreadsheet updates were needed.")

    print("\nProcessing complete!")

if __name__ == '__main__':
    main()