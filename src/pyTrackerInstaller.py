import json
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font

# Constants
CONFIG_DIR = "config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "email_config.json")
THEME_COLOR = "#3498db"  # A nice blue color
BG_COLOR = "#f5f5f5"     # Light gray background
BTN_COLOR = "#2980b9"    # Darker blue for buttons

class SetupWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Window setup
        self.title("pyTracker Setup Wizard")
        self.geometry("600x450")
        self.configure(bg=BG_COLOR)
        self.resizable(False, False)
        
        # Center window on screen
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')
        
        # Add an icon (replace with your own .ico file)
        # self.iconbitmap("icon.ico")
        
        self.config_data = {}
        
        # Add a header
        self.header_frame = tk.Frame(self, bg=THEME_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Arial", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Setup", 
            font=header_font, 
            bg=THEME_COLOR, 
            fg="white"
        ).pack(pady=20)
        
        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(padx=40, pady=20, fill="both", expand=True)
        
        # Setup title and instructions
        title_font = Font(family="Arial", size=14, weight="bold")
        label_font = Font(family="Arial", size=10)
        
        tk.Label(
            self.main_frame, 
            text="Welcome to pyTracker", 
            font=title_font, 
            bg=BG_COLOR
        ).pack(anchor="w", pady=(0, 5))
        
        tk.Label(
            self.main_frame, 
            text="Please configure the following settings to get started:", 
            font=label_font,
            bg=BG_COLOR
        ).pack(anchor="w", pady=(0, 20))
        
        # Input frames with improved styling
        # Email
        self.create_input_field("Email Address:", "email_entry")
        
        # App Key
        self.create_input_field("App Key:", "appKey_entry", show="*")
        
        #Sheets ID
        self.create_input_field("Sheet ID:", "sheetID_entry")
        
        # Output Directory
        self.dir_frame = ttk.Frame(self.main_frame)
        self.dir_frame.pack(fill="x", pady=10)
        
        ttk.Label(
            self.dir_frame, 
            text="Output Directory:", 
            width=15, 
            anchor="w"
        ).pack(side="left")
        
        self.dir_entry = ttk.Entry(self.dir_frame)
        self.dir_entry.pack(side="left", fill="x", expand=True, padx=(5, 5))
        
        # Style the browse button
        browse_style = ttk.Style()
        browse_style.configure("Browse.TButton", background=THEME_COLOR)
        
        self.browse_button = ttk.Button(
            self.dir_frame, 
            text="Browse...", 
            command=self.browse_directory,
            style="Browse.TButton"
        )
        self.browse_button.pack(side="left", padx=5)
        
        # Separator before buttons
        ttk.Separator(self.main_frame, orient='horizontal').pack(fill='x', pady=20)
        
        # Button frame with modern styling
        self.button_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        self.button_frame.pack(fill="x", pady=10)
        
        # Style the buttons
        button_style = ttk.Style()
        button_style.configure("TButton", font=("Arial", 10))
        button_style.configure("Primary.TButton", background=BTN_COLOR)
        
        self.cancel_button = ttk.Button(
            self.button_frame, 
            text="Cancel", 
            command=self.destroy
        )
        self.cancel_button.pack(side="right", padx=5)
        
        self.save_button = ttk.Button(
            self.button_frame, 
            text="Save Configuration", 
            command=self.save_config,
            style="Primary.TButton"
        )
        self.save_button.pack(side="right", padx=5)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set("Enter your information to begin")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground="gray"
        )
        self.status_label.pack(side="bottom", pady=10)
        
    def create_input_field(self, label_text, entry_name, show=None):
        """Helper method to create standardized input fields"""
        frame = ttk.Frame(self.main_frame)
        frame.pack(fill="x", pady=10)
        
        ttk.Label(
            frame, 
            text=label_text, 
            width=15, 
            anchor="w"
        ).pack(side="left")
        
        entry = ttk.Entry(frame, show=show)
        entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Store the entry widget as an attribute
        setattr(self, entry_name, entry)
        
    def browse_directory(self):
        # Ensure the config directory exists
        if not os.path.exists(CONFIG_DIR):
            os.makedirs(CONFIG_DIR)
        
        # Open the directory dialog with the default path set to the config directory
        directory = filedialog.askdirectory(initialdir=CONFIG_DIR)
        if directory:
            self.dir_entry.delete(0, tk.END)
            self.dir_entry.insert(0, directory)
            self.status_var.set(f"Directory selected: {directory}")
            
    def save_config(self):
        # Get values from entries
        self.config_data["email"] = self.email_entry.get()
        self.config_data["appKey"] = self.appKey_entry.get()
        self.config_data["output_dir"] = self.dir_entry.get()
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
            messagebox.showinfo("Success", "Configuration saved successfully!")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
            self.status_var.set("Error saving configuration")


def load_config():
    """Load the configuration file if it exists or run the setup wizard."""
    # Check if config file exists and try to load it
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config: {e}")
            # Continue to setup wizard if an error occurs
    
    # If we get here, either the file doesn't exist or loading failed
    # Create and run the setup wizard
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
    
    # After the wizard closes, check if config file was created
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    else:
        # User cancelled setup
        return None
        
        
def main():
    # Load or create configuration
    config = load_config()
    
    if config is None:
        print("Setup was cancelled. Exiting application.")
        return
    
    # Your main application logic here
    print(f"Running with email: {config['email']}")
    print(f"Output directory: {config['output_dir']}")
    
    # Start your actual application...
    # app = YourMainApplication(config)
    # app.run()


if __name__ == "__main__":
    main()