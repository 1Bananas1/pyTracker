import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
import os
import json
import sys

CONFIG_DIR = "config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "email_config.json")
THEME_COLOR = "#3498db"  # A nice blue color
BG_COLOR = "#f5f5f5"     # Light gray background
BTN_COLOR = "#2980b9"    # Darker blue for buttons


class SetupWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("pyTracker Setup Wizard")
        self.geometry("600x450")
        self.configure(bg=BG_COLOR)
        self.resizable(False, False)
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../public/icon.ico"))
        self.iconbitmap(icon_path)

        # Custom font
        self.custom_font = Font(family="Helvetica", size=12)

        # Create and configure widgets
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'+{x}+{y}')

        if os.path.exists(CONFIG_FILE):
            self.load_main_screen()
        else:
            self.load_setup_screen()

    def create_input_field(self, label_text, entry_name, show=None, default_value=""):
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
        
        # Set default value if provided
        if default_value:
            entry.insert(0, default_value)
        
        # Store the entry widget as an attribute
        setattr(self, entry_name, entry)

    def create_directory_field(self, label_text, entry_name, default_value=""):
        """Helper method to create a directory input field with a browse button"""
        frame = ttk.Frame(self.main_frame)
        frame.pack(fill="x", pady=10)
        
        ttk.Label(
            frame, 
            text=label_text, 
            width=15, 
            anchor="w"
        ).pack(side="left")
        
        entry = ttk.Entry(frame)
        entry.pack(side="left", fill="x", expand=True, padx=(5, 0))
        
        # Set default value if provided
        if default_value:
            entry.insert(0, default_value)
        
        browse_button = ttk.Button(frame, text="Browse...", command=lambda: self.browse_directory(entry))
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
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
            self.status_var.set("Error saving configuration")

    def load_setup_screen(self):
        self.title('pyTracker Initial Setup')
        for widget in self.winfo_children():
            widget.destroy()

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
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Add input fields
        self.create_input_field("Email Address:", "email_entry")
        self.create_input_field("App Key:", "appKey_entry", show="*")
        self.create_input_field("Sheet ID:", "sheetID_entry")
        self.create_directory_field("Output Directory:", "output_dir_entry")

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

    def load_main_screen(self):
        self.title('pyTracker Modify Setup')
        for widget in self.winfo_children():
            widget.destroy()

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


def main():
    config = loadConfig()


if __name__ == "__main__":
    main()