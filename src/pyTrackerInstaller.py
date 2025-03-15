import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os
import json
import sys, time
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread,pyglet

CONFIG_DIR = "config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "email_config.json")

BG_COLOR = '#010104'     # Light gray background
PRIMARY_COLOR = "#386c54"  # Dark green for primary elements
ACCENT_COLOR = "#4A90E2"  # Accent color for highlights
SECONDARY_COLOR = "#217346"  # Secondary color for less important elements
TEXT_COLOR = "#ebe9fc"  # Text color for readability

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


class SetupWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("pyTracker Setup Wizard")
        self.geometry("600x450")
        self.configure(bg=BG_COLOR)
        self.resizable(False, False)
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../public/icons/icon.ico"))
        self.iconbitmap(icon_path)

        # Initialize config_data
        self.config_data = {}

        # Custom font
        self.custom_font = Font(family="Inter Regular", size=12)

        # Create and configure widgets
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')

        # Configure styles
        self.configure_styles()

        if os.path.exists(CONFIG_FILE):
            self.load_main_screen()
        else:
            self.load_setup_screen()

    def configure_styles(self):
        """Configure styles for ttk widgets"""
        button_style = ttk.Style()
        button_style.configure("TButton", font=("Inter Regular", 10), background=ACCENT_COLOR, foreground=TEXT_COLOR)
        button_style.map("TButton", background=[('active', SECONDARY_COLOR)])
        
        # Configure Primary.TButton style
        button_style.configure("Primary.TButton", background=ACCENT_COLOR)
        button_style.map("Primary.TButton", background=[('active', SECONDARY_COLOR)])
        
        # Configure entry style with correct background
        entry_style = ttk.Style()
        entry_style.configure("TEntry", 
                            borderwidth=0, 
                            relief="flat", 
                            fieldbackground=BG_COLOR,  # This sets the actual text area background
                            background=BG_COLOR,
                            foreground=TEXT_COLOR)
        
        # You may also need to map the background color for different states
        entry_style.map("TEntry",
                    fieldbackground=[('readonly', BG_COLOR), ('disabled', BG_COLOR)],
                    background=[('readonly', BG_COLOR), ('disabled', BG_COLOR)])

    def create_input_field(self, label_text, entry_name, show=None, default_value=""):
        """Helper method to create standardized input fields"""
        frame = tk.Frame(self.main_frame, bg=BG_COLOR)  # Use tk.Frame with bg color
        frame.pack(fill="x", pady=2)
        
        # Use tk.Label instead of ttk.Label for better styling control
        tk.Label(
            frame, 
            text=label_text, 
            width=15, 
            bg=SECONDARY_COLOR,  # Use bg instead of background for tk widgets
            fg=TEXT_COLOR,       # Use fg instead of foreground for tk widgets
            anchor="w"
        ).pack(side="left", padx=(0, 0))  # Remove extra padding
        
        # Replace ttk.Entry with tk.Entry
        entry = tk.Entry(
            frame,
            show=show,
            background=BG_COLOR,  # Direct background setting
            foreground=TEXT_COLOR,  # Direct text color
            insertbackground=TEXT_COLOR,  # Cursor color
            relief="flat",
            highlightthickness=1,
            highlightbackground=SECONDARY_COLOR,
            highlightcolor=ACCENT_COLOR
        )
        entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Set default value if provided
        if default_value:
            entry.insert(0, default_value)
        
        # Store the entry widget as an attribute
        setattr(self, entry_name, entry)

    def create_directory_field(self, label_text, entry_name, default_value=""):
        """Helper method to create a directory input field with a browse button"""
        frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        frame.pack(fill="x", pady=2)
        
        # Use tk.Label with bg parameter
        tk.Label(
            frame, 
            text=label_text, 
            width=15, 
            bg=SECONDARY_COLOR,
            fg=TEXT_COLOR,
            anchor="w"
        ).pack(side="left", padx=(0, 0))
        
        entry = tk.Entry(
            frame,
            background=BG_COLOR,
            foreground=TEXT_COLOR,
            insertbackground=TEXT_COLOR,
            relief="flat",
            highlightthickness=1,
            highlightbackground=SECONDARY_COLOR,
            highlightcolor=ACCENT_COLOR
        )
        entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Set default value if provided
        if default_value:
            entry.insert(0, default_value)
        
        # Switch to tk.Button for direct styling control
        browse_button = tk.Button(
            frame, 
            text="Browse...", 
            command=lambda: self.browse_directory(entry),
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR,
            padx=10
        )
        browse_button.pack(side="left", padx=(5, 0))
        
        # Store the entry widget as an attribute
        setattr(self, entry_name, entry)

    def browse_directory(self, entry):
        """Open a directory selection dialog and set the selected directory in the entry"""
        directory = filedialog.askdirectory()
        if directory:
            entry.delete(0, tk.END)
            entry.insert(0, directory)

    def save_config(self):
        # Get values from entries
        self.config_data["email"] = self.email_entry.get()
        self.config_data["appKey"] = self.appKey_entry.get()
        self.config_data["output_dir"] = self.output_dir_entry.get()
        self.config_data["sheetID"] = self.sheetID_entry.get()
        
        # Validate inputs
        if not self.config_data["email"]:
            messagebox.showerror("Error", "Email cannot be empty!")
            self.status_var.set("Error: Email field required")
            return
        
        if not self.config_data["appKey"]:
            messagebox.showerror("Error", "App Key cannot be empty!")
            self.status_var.set("Error: App Key field required")
            return
        
        if not self.config_data["sheetID"]:
            messagebox.showerror("Error", "Sheet ID cannot be empty!")
            self.status_var.set("Error: Sheet ID field required")
            return
        
        # Create output directory if it doesn't exist
        if self.config_data["output_dir"] and not os.path.exists(self.config_data["output_dir"]):
            try:
                os.makedirs(self.config_data["output_dir"])
            except Exception as e:
                messagebox.showerror("Error", f"Could not create directory: {str(e)}")
                self.status_var.set("Error creating directory")
                return
        
        # Ensure config directory exists
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        
        # Save configuration
        try:
            with open(CONFIG_FILE, "w") as f:
                json.dump(self.config_data, f)
            messagebox.showinfo("Success", "Configuration saved successfully! Please download your credentials file.")
            self.sheetSetup()  # Call sheetSetup after saving the configuration
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
            self.status_var.set("Error saving configuration")

    def load_setup_screen(self):
        self.title('pyTracker Initial Setup')
        for widget in self.winfo_children():
            widget.destroy()

        # Add a header
        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Setup", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)

        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Add input fields
        self.create_input_field("Email Address:", "email_entry")
        self.create_input_field("App Key:", "appKey_entry", show="*")
        self.create_input_field("Sheet ID:", "sheetID_entry")
        self.create_directory_field("Output Directory:", "output_dir_entry")

        separator = tk.Frame(self.main_frame, height=1, bg=SECONDARY_COLOR)
        separator.pack(fill='x', pady=20)
        
        # Button frame with modern styling
        self.button_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        self.button_frame.pack(fill="x", pady=10)
        
        # Style the buttons
        button_style = ttk.Style()
        button_style.configure("TButton", font=("Inter Regular", 10))
        button_style.configure("Primary.TButton", background=SECONDARY_COLOR)
        
        self.cancel_button = tk.Button(
            self.button_frame, 
            text="Cancel", 
            command=self.destroy,
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            padx=10,
            pady=5,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.cancel_button.pack(side="right", padx=5)

        self.save_button = tk.Button(
            self.button_frame, 
            text="Save Configuration", 
            command=self.save_config,
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            padx=10,
            pady=5,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.save_button.pack(side="right", padx=5)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set("Enter your information to begin")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)

    def load_main_screen(self):
        self.title('pyTracker Modify Setup')
        for widget in self.winfo_children():
            widget.destroy()

        # Add a header
        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Setup", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)

        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        continuebox = messagebox.askokcancel("pyTracker Setup", "An existing config file was found. Would you like to modify it?")
        if not continuebox:
            self.destroy()
            return

        with open(CONFIG_FILE, 'r') as f:
            self.config_data = json.load(f)

        # Add input fields with old values
        self.create_input_field("Email Address:", "email_entry", default_value=self.config_data.get("email", ""))
        self.create_input_field("App Key:", "appKey_entry", show="*", default_value=self.config_data.get("appKey", ""))
        self.create_input_field("Sheet ID:", "sheetID_entry", default_value=self.config_data.get("sheetID", ""))
        self.create_directory_field("Output Directory:", "output_dir_entry", default_value=self.config_data.get("output_dir", ""))

        separator = tk.Frame(self.main_frame, height=1, bg=SECONDARY_COLOR)
        separator.pack(fill='x', pady=20)
        
        # Button frame with modern styling
        self.button_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        self.button_frame.pack(fill="x", pady=10)
        
        # Style the buttons
        button_style = ttk.Style()
        button_style.configure("TButton", font=("Inter Regular", 10))
        button_style.configure("Primary.TButton", background=SECONDARY_COLOR)
        
        self.cancel_button = tk.Button(
            self.button_frame, 
            text="Cancel", 
            command=self.destroy,
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            padx=10,
            pady=5,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.cancel_button.pack(side="right", padx=5)

        self.save_button = tk.Button(
            self.button_frame, 
            text="Save Configuration", 
            command=self.save_config,
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            padx=10,
            pady=5,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.save_button.pack(side="right", padx=5)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set("Enter your information to begin")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)

    def sheetSetup(self):
        self.title('pyTracker Sheet Setup')
        for widget in self.winfo_children():
            widget.destroy()

        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Setup", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)

        setupSheet = messagebox.askokcancel("pyTracker Setup", "Would you like to setup your Google Sheet now?")
        if not setupSheet:
            self.destroy()
            return
        
        if not os.path.exists(CONFIG_FILE):
            messagebox.showerror("Error", "No configuration file found!")
            self.destroy()
            return
        
        if not os.path.exists(os.path.join(CONFIG_DIR, "credentials.json")):
            messagebox.showerror("Error", "No credentials file found!")
            self.destroy()
            return
        
        print("Starting Google Sheet setup...")

        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
            
        SPREADSHEET_ID = config['sheetID']
        creds = get_credentials()
        service = build('sheets', 'v4', credentials=creds)
        gc = gspread.authorize(creds)
        was_initialized = initialize_sheets(service, SPREADSHEET_ID, gc)
        if was_initialized:
            print("First-time setup completed successfully.")
            
            # Create a success frame
            success_frame = tk.Frame(self, bg=BG_COLOR)
            success_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
            
            # Success message with checkmark
            success_message = tk.Label(
                success_frame,
                text="✓ First-time setup completed successfully!",
                font=("Inter Regular", 14),
                fg=TEXT_COLOR,
                bg=BG_COLOR
            )
            success_message.pack(pady=(50, 0))
            
            # Update the window to show the success message
            self.update()
            
            # Wait 2 seconds and then close
            self.after(2000, self.destroy)
        else:
            print("Using existing sheet structure.")
            messagebox.showinfo("Setup Complete", "Using existing sheet structure.")
            self.destroy()
        

def loadConfig():
    app = SetupWizard()
    
    # Set theme for ttk widgets
    style = ttk.Style(app)
    try:
        if os.name == 'nt':  # Windows
            style.theme_use('vista')
        else:  # macOS or Linux
            style.theme_use('clam')
    except tk.TclError:
        # Fall back to a default theme if the specified one isn't available
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
    
    # Run the application
    app.mainloop()


def get_credentials():
    with open(CONFIG_FILE, 'r') as f:
        config = json.load(f)

    OUTPUT_DIR = config['output_dir']
    creds = None
    # The file token.json stores the user's access and refresh tokens
    token_path = os.path.join(CONFIG_DIR, 'token.json')
    
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

def check_if_initialized(gc_client, spreadsheet_id):
    """Check if this spreadsheet has already been set up with our custom sheets."""
    spreadsheet = gc_client.open_by_key(spreadsheet_id)
    worksheet_titles = [ws.title for ws in spreadsheet.worksheets()]
    
    # If both Applications and Backend sheets exist, consider it initialized
    return 'Applications' in worksheet_titles and 'Backend' in worksheet_titles

def initialize_sheets(service, spreadsheet_id, gc_client):
    """Set up the spreadsheet with Applications and Backend sheets."""
    is_initialized = check_if_initialized(gc_client, spreadsheet_id)
    
    if (is_initialized):
        print("Spreadsheet already initialized with required sheets.")
        return False
    
    print("Initializing spreadsheet with required sheets...")
    spreadsheet = gc_client.open_by_key(spreadsheet_id)
    
    # Step 1: Create new sheets first
    # Create Applications sheet if it doesn't exist
    if 'Applications' not in [ws.title for ws in spreadsheet.worksheets()]:
        spreadsheet.add_worksheet(title='Applications', rows=1000, cols=20)
        print("Created 'Applications' sheet")
    
    # Create Backend sheet if it doesn't exist
    if 'Backend' not in [ws.title for ws in spreadsheet.worksheets()]:
        spreadsheet.add_worksheet(title='Backend', rows=1000, cols=20)
        print("Created 'Backend' sheet")
    
    # Step 2: Set up headers in Applications sheet if needed
    applications_sheet = spreadsheet.worksheet('Applications')
    if not applications_sheet.cell(1, 1).value:
        applications_sheet.update('A1:F1', [['Status', 'Company', 'Date Applied', 'Last Updated','Link','Role']])
        applications_sheet.freeze(rows=1)
        applications_sheet.format('A1:F1', {'textFormat': {'bold': True}})
        applications_sheet.format('A1:F1', {"backgroundColor": {"red": 0.961, "green": 0.961, "blue": 0.961}})
        print("Initialized headers in Applications sheet")

    backend_sheet = spreadsheet.worksheet('Backend')
    if not backend_sheet.cell(1, 1).value:
        backend_sheet.update('A1:C1', [['Company', 'Company ID', 'Role']])
        backend_sheet.freeze(rows=1)
        backend_sheet.format('A1:C1', {'textFormat': {'bold': True}})
        backend_sheet.format('A1:C1', {"backgroundColor": {"red": 0.961, "green": 0.961, "blue": 0.961}})
        print("Initialized headers in Backend sheet")
    
    # Step 3: Now that we have our custom sheets, remove default sheets
    remove_default_sheets(service, spreadsheet_id, gc_client)
    
    return True

def remove_default_sheets(service, spreadsheet_id, gc_client):
    """Remove default sheets like Sheet1, Sheet2, etc. after ensuring we have custom sheets."""
    # Open the spreadsheet using gspread
    spreadsheet = gc_client.open_by_key(spreadsheet_id)
    
    # Get all worksheets
    worksheets = spreadsheet.worksheets()
    
    # Default sheet names to remove
    default_names = ["Sheet1", "Sheet2", "Sheet3"]
    
    # Track if we've removed sheets 
    sheets_removed = False
    
    # Remove default sheets - now safe to do since we've created our custom sheets
    for worksheet in worksheets:
        if worksheet.title in default_names:
            try:
                print(f"Deleting default worksheet: {worksheet.title}")
                spreadsheet.del_worksheet(worksheet)
                sheets_removed = True
            except Exception as e:
                print(f"Could not delete {worksheet.title}: {e}")
    
    return sheets_removed

def main():
    pyglet.font.add_file('public/fonts/Inter-VariableFont_opsz,wght.ttf')
    config = loadConfig()


if __name__ == "__main__":
    main()