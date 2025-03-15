
# pyTracker

A way for computer science student to keep track of their internships


## Installation
This is a slight bit of a process so please follow ALL steps

#### App Key
Make a Google [App Key](https://myaccount.google.com/u/1/apppasswords) to use during setup.

#### Credentials File

  1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
  2. Create a new project (or select an existing one)
  3. Add your email as a dev tester
  4. Enable the Gmail and Google Sheets API
  5. Go to "APIs & Services" > "Credentials"
  6. Create OAuth 2.0 Client ID credentials
  7. Download the JSON file, which will be your credentials file
  8. Place credentials file into your config with the name `credentials.json`

### Initial setup

```bash
  py -m venv .venv
  .venv\Scripts\activate
  pip install requirements.txt
```

```
  python ./src/pyTrackerInstaller2.py
```

## Installer
Make a new google sheet, and take the ID found with `https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID_HERE/edit#gid=0`

Enter the required details, and press set up google sheet

Download Ollama
    