import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import os, requests
from PIL import Image, ImageTk
from io import BytesIO
import json, GPUtil, pandas as pd
import sys, time
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread, pyglet
from tkinter import PhotoImage


# Create a rounded rectangle method for tk.Canvas
tk.Canvas.create_rounded_rectangle = lambda self, x1, y1, x2, y2, radius=25, **kwargs: \
    self.create_polygon(
        x1+radius, y1,
        x2-radius, y1,
        x2, y1,
        x2, y1+radius,
        x2, y2-radius,
        x2, y2,
        x2-radius, y2,
        x1+radius, y2,
        x1, y2,
        x1, y2-radius,
        x1, y1+radius,
        x1, y1,
        smooth=True, **kwargs)

CONFIG_DIR = "config"
CONFIG_FILE = os.path.join(CONFIG_DIR, "email_config.json")

BG_COLOR = '#010104'     # Light gray background
PRIMARY_COLOR = "#386c54"  # Dark green for primary elements
ACCENT_COLOR = "#4A90E2"  # Accent color for highlights
SECONDARY_COLOR = "#217346"  # Secondary color for less important elements
TEXT_COLOR = "#ebe9fc"  # Text color for readability

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def load_image_from_url(url, width=None, height=None):
    response = requests.get(url)
    img_data = BytesIO(response.content)
    img = Image.open(img_data)
    
    if width and height:
        img = img.resize((width, height), Image.LANCZOS)
    
    return ImageTk.PhotoImage(img)


class SetupWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("pyTracker Setup Wizard")
        self.geometry("600x450")
        self.configure(bg=BG_COLOR)
        self.resizable(False, False)
        icon_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../public/icons/icon.ico"))
        try:
            self.iconbitmap(icon_path)
        except tk.TclError:
            print(f"Warning: Could not load icon from {icon_path}")

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
        
        # Initialize selection tracking
        self.finalSelection = []
        
        # Check if configuration exists and load appropriate screen
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
            # If user chooses to skip sheet setup, go to model selection
            self.getModel()
            return
        
        if not os.path.exists(CONFIG_FILE):
            messagebox.showerror("Error", "No configuration file found!")
            # Even after error, go to model selection
            self.getModel()
            return
        
        if not os.path.exists(os.path.join(CONFIG_DIR, "credentials.json")):
            messagebox.showerror("Error", "No credentials file found!")
            # Even after error, go to model selection
            self.getModel()
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
        else:
            print("Using existing sheet structure.")
        
        # After completing sheet setup, proceed to model selection
        self.getModel()

    def getBestModel(self):
        GPUs = GPUtil.getGPUs()
        totalFree = 0
        for gpu in GPUs:
            totalFree += gpu.memoryTotal
        
        try:
            gpuList = pd.read_csv('public/data/ollama_nlp_models.csv')
            recommendedFiltered = gpuList.loc[gpuList['recVRAM'] <= totalFree]
            minFiltered = gpuList.loc[gpuList['VRAM'] <= totalFree]
            
            # Get the index of the row with max Parameter Size, then select the entire row
            if not minFiltered.empty:
                max_idx = minFiltered['Parameter Size'].idxmax()
                selected = minFiltered.loc[max_idx]
                model_version = selected['Model Name'] + ':' + selected['normP']
                
                # Save the model version to the configuration file
                try:
                    # Read existing config if available
                    config_data = {}
                    if os.path.exists(CONFIG_FILE):
                        with open(CONFIG_FILE, 'r') as f:
                            config_data = json.load(f)
                    
                    # Add or update the model information
                    config_data['model_version'] = model_version
                    
                    # Ensure config directory exists
                    if not os.path.exists(CONFIG_DIR):
                        os.makedirs(CONFIG_DIR)
                    
                    # Save the updated configuration
                    with open(CONFIG_FILE, 'w') as f:
                        json.dump(config_data, f)
                    
                    self.status_var.set(f"Selected model: {model_version}")
                    messagebox.showinfo("Model Selected", f"The model {model_version} has been selected and saved to your configuration.")
                    
                    # Close the application after successfully setting model
                    self.destroy()
                    
                    return selected
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save model selection: {str(e)}")
                    return None
            else:
                messagebox.showwarning("No Models Found", "No compatible models were found for your GPU.")
                print("No compatible models found")
                return None
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting best model: {str(e)}")
            return None
        
    def flexibleCommand(self, **kwargs):
        pass
    
    def getModel(self):
        self.title('pyTracker Model Selection')
        for widget in self.winfo_children():
            widget.destroy()

        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Model Selection", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)

        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Center container for buttons
        center_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        center_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Button styles - large with accent color
        button_font = Font(family="Inter Regular", size=14, weight="bold")
        button_width = 20
        button_height = 3

        # Manual Setup Button
        self.manual_button = tk.Button(
            center_frame, 
            text="Manual Setup", 
            command=self.manualModel,
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            font=button_font,
            relief="flat",
            borderwidth=0,
            width=button_width,
            height=button_height,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.manual_button.pack(side="left", padx=15)

        # Automatic Setup Button
        self.automatic_button = tk.Button(
            center_frame, 
            text="Automatic", 
            command=self.getBestModel,
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            font=button_font,
            relief="flat",
            borderwidth=0,
            width=button_width,
            height=button_height,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR
        )
        self.automatic_button.pack(side="right", padx=15)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set("Select setup method")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)

    def select_model(self, model_id):
        """
        Load the model data from CSV and filter by model author/family.
        
        Args:
            model_id: The model identifier to filter by (e.g., "llama", "qwen")
            
        Returns:
            DataFrame containing filtered model data
        """
        try:
            models_df = pd.read_csv("public/data/ollama_nlp_models.csv")
            filtered_models = models_df[models_df["modelAuthor"] == model_id]
            return filtered_models
        except Exception as e:
            print(f"Error loading model data: {e}")
            messagebox.showerror("Error", f"Failed to load model data: {e}")
            return pd.DataFrame()  # Return empty DataFrame on error

    def gotoModelVersionSelection(self, model):
        """
        Display model version selection screen after a model family is selected
        
        Args:
            model: Selected model family (e.g., "llama", "qwen")
        """
        self.title('pyTracker Model Selection')
        for widget in self.winfo_children():
            widget.destroy()
        
        # Store the selected model family
        self.finalSelection = [model]
        
        # Get filtered DataFrame for this model
        filtered = self.select_model(model)
        if filtered.empty:
            messagebox.showerror("Error", f"No versions found for {model}")
            self.manualModel()  # Go back to model selection
            return
        
        # Create header
        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text=f"Select which {model} version to use with pyTracker:", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)
        
        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Create frame for version cards
        models_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        models_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Loop through unique versions and create cards
        versions = filtered['version'].unique()
        for i, version in enumerate(versions):
            # Find the corresponding row for this version
            matching_rows = filtered[filtered['version'] == version]
            if not matching_rows.empty:
                matching_row = matching_rows.iloc[0]
                img_path = matching_row['imgloc']
                
                # Create card with model logo
                card = self.createVersionCard(models_frame, str(version), img_path)
                card.grid(row=i//3, column=i%3, padx=10, pady=10)
        
        # Add back button
        back_button = tk.Button(
            self.main_frame,
            text="← Back",
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            command=self.manualModel
        )
        back_button.pack(side="bottom", pady=10)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set(f"Select a {model} version")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)

    def createVersionCard(self, parent, version, img_path):
        """
        Create a model version selection card
        
        Args:
            parent: Parent widget
            version: Version number as string
            img_path: Path to the model image
            
        Returns:
            Frame containing the version card
        """
        # Create container frame
        container = tk.Frame(
            parent,
            bg=BG_COLOR,
            width=120,
            height=150,
            bd=0
        )
        container.pack_propagate(False)
        
        # Create a canvas with rounded border
        canvas = tk.Canvas(
            container,
            bg=BG_COLOR,
            highlightthickness=0,
            width=100,
            height=100
        )
        canvas.pack(pady=(0, 5))
        
        # Draw rounded rectangle for border
        border_id = canvas.create_rounded_rectangle(
            5, 5, 95, 95,  # Coordinates
            radius=10,     # Radius for rounded corners
            fill=BG_COLOR,
            outline=PRIMARY_COLOR,
            width=2
        )
        
        # Load and display model image
        try:
            img = Image.open(img_path)
            img = img.resize((70, 70), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            
            # Create larger image for hover state
            hover_img = Image.open(img_path)
            hover_img = hover_img.resize((80, 80), Image.LANCZOS)
            hover_photo = ImageTk.PhotoImage(hover_img)
            
            # Add image to canvas
            image_id = canvas.create_image(50, 50, image=photo, anchor="center")
            canvas.image = photo  # Keep reference
            canvas.hover_image = hover_photo  # Keep reference
            
            # Add hover effects
            def on_enter(e, canv=canvas, h_img=hover_photo, b_id=border_id):
                canv.itemconfig(b_id, outline=ACCENT_COLOR, width=3)
                canv.itemconfig(canv.find_withtag("current")[0], image=h_img)
                
            def on_leave(e, canv=canvas, o_img=photo, b_id=border_id):
                canv.itemconfig(b_id, outline=PRIMARY_COLOR, width=2)
                canv.itemconfig(canv.find_withtag("current")[0], image=o_img)
                
            canvas.tag_bind(image_id, "<Enter>", on_enter)
            canvas.tag_bind(image_id, "<Leave>", on_leave)
            canvas.tag_bind(image_id, "<Button-1>", 
                        lambda e, v=version: self.gotoModelParameterSelection(v))
            
        except Exception as e:
            print(f"Error loading image for version {version}: {e}")
            # Create a fallback text representation if image fails to load
            canvas.create_text(50, 50, text=version, fill=TEXT_COLOR, font=("Inter Regular", 18))
        
        # Version label
        tk.Label(
            container,
            text=version,
            fg=TEXT_COLOR,
            bg=BG_COLOR,
            font=("Inter Regular", 14)
        ).pack(pady=5)
        
        return container

    def gotoModelParameterSelection(self, modelVersion):
        """
        Handle selection of a specific model parameter when a model version card is clicked.
        Will display parameter options for the selected model version and save the selection.
        
        Args:
            modelVersion: The version of the model that was selected (e.g., "2.0", "3.1")
        """
        self.title('pyTracker Model Parameter Selection')
        for widget in self.winfo_children():
            widget.destroy()
        
        # Get the filtered models for this version
        filtered_df = self.select_model(self.finalSelection[0])  # Get base model (e.g., "llama")
        version_df = filtered_df[filtered_df['version'] == float(modelVersion)]  # Filter by version number
        
        # Create header
        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text=f"Select {self.finalSelection[0]} {modelVersion} Parameter Size", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)
        
        # Main content frame
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Instructions label
        tk.Label(
            self.main_frame,
            text=f"Select which parameter size to use:",
            font=("Inter Regular", 12),
            bg=BG_COLOR,
            fg=TEXT_COLOR
        ).pack(pady=(0, 10))
        
        # Create frame for parameter cards
        params_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        params_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Get unique parameter sizes for this version
        param_sizes = version_df['normP'].unique()
        
        # Create a card for each parameter size
        for i, param_size in enumerate(param_sizes):
            param_df = version_df[version_df['normP'] == param_size]
            if not param_df.empty:
                card = self.createParameterCard(params_frame, param_size, param_df.iloc[0], modelVersion)
                card.grid(row=i//3, column=i%3, padx=10, pady=10)
        
        # Add back button
        back_button = tk.Button(
            self.main_frame,
            text="← Back",
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            command=lambda: self.gotoModelVersionSelection(self.finalSelection[0])
        )
        back_button.pack(side="bottom", pady=10)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set(f"Select a parameter size for {self.finalSelection[0]} {modelVersion}")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)

    def createParameterCard(self, parent, param_size, row_data, model_version):
        """
        Create a card for a specific parameter size option
        
        Args:
            parent: Parent widget
            param_size: The parameter size to display (e.g., "7B", "13B")
            row_data: DataFrame row containing model information
            model_version: The version of the model
            
        Returns:
            Frame containing the parameter card
        """
        # Get the required VRAM
        vram_required = row_data['VRAM']
        
        # Create container frame
        container = tk.Frame(
            parent,
            bg=BG_COLOR,
            width=150,
            height=180,
            bd=0,
            highlightbackground=PRIMARY_COLOR,
            highlightthickness=2,
            highlightcolor=ACCENT_COLOR
        )
        container.pack_propagate(False)
        
        # Parameter size label
        param_label = tk.Label(
            container,
            text=param_size,
            font=("Inter Regular", 16, "bold"),
            bg=BG_COLOR,
            fg=TEXT_COLOR
        )
        param_label.pack(pady=(20, 5))
        
        # Parameter details
        details_text = f"Parameters: {row_data['RAM Requirement (GB)']}\nVRAM: {vram_required} GB"
        details_label = tk.Label(
            container,
            text=details_text,
            font=("Inter Regular", 10),
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            justify=tk.LEFT
        )
        details_label.pack(pady=5)
        
        # Select button
        select_btn = tk.Button(
            container,
            text="Select",
            bg=ACCENT_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=0,
            activebackground=SECONDARY_COLOR,
            activeforeground=TEXT_COLOR,
            command=lambda p=param_size: self.saveModelSelection(self.finalSelection[0], model_version, p)
        )
        select_btn.pack(pady=10)
        
        # Add hover effects
        def on_enter(e, btn=select_btn):
            btn.config(bg=SECONDARY_COLOR)
        
        def on_leave(e, btn=select_btn):
            btn.config(bg=ACCENT_COLOR)
        
        select_btn.bind("<Enter>", on_enter)
        select_btn.bind("<Leave>", on_leave)
        
        return container

    def saveModelSelection(self, model_family, version, param_size):
        """
        Save the selected model configuration to the config file
        
        Args:
            model_family: Model family (e.g., "llama", "qwen")
            version: Model version (e.g., "2.0", "3.1")
            param_size: Parameter size (e.g., "7b", "13b")
        """
        try:
            # Format the model string as expected by Ollama
            # Ollama model formats follow patterns like: llama3.1, llama3.1:8b, gemma3:27b, qwen:72b, phi4-mini
            
            # First normalize the version (remove decimal if it's .0)
            if version.endswith(".0"):
                version = version[:-2]
                
            # Format based on model family conventions
            if model_family == "mistral":
                # Special case for mistral which doesn't use version number
                if param_size.lower() == "default" or not param_size:
                    model_string = f"{model_family}"
                else:
                    model_string = f"{model_family}:{param_size.lower()}"
            elif model_family == "llama":
                # Llama uses format like llama3.1 or llama3.1:8b
                if param_size.lower() == "default" or not param_size:
                    model_string = f"{model_family}{version}"
                else:
                    model_string = f"{model_family}{version}:{param_size.lower()}"
            elif model_family in ["gemma", "qwen", "mistral", "phi"]:
                # Most models use format like gemma3:27b or phi4-mini
                if param_size.lower() == "default" or not param_size:
                    model_string = f"{model_family}{version}"
                else:
                    # Handle special cases like phi4-mini
                    if param_size.lower() == "mini":
                        model_string = f"{model_family}{version}-{param_size.lower()}"
                    else:
                        model_string = f"{model_family}{version}:{param_size.lower()}"
            else:
                # Default format for other models
                if param_size.lower() == "default" or not param_size:
                    model_string = f"{model_family}{version}"
                else:
                    model_string = f"{model_family}{version}:{param_size.lower()}"
            
            # Load existing config
            config_data = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config_data = json.load(f)
            
            # Update model version
            config_data['model_version'] = model_string
            
            # Save config
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f)
            
            # Show success message
            self.status_var.set(f"Model {model_string} saved to configuration!")
            messagebox.showinfo("Success", f"Selected model: {model_string}")
            
            # Close the window or navigate to another screen
            self.destroy()
            
        except Exception as e:
            self.status_var.set(f"Error saving model selection: {str(e)}")
            messagebox.showerror("Error", f"Failed to save model selection: {str(e)}")
            
    def manualModel(self):
        """
        Display the initial model family selection screen with model cards
        """
        self.title('pyTracker Model Selection')
        for widget in self.winfo_children():
            widget.destroy()

        self.header_frame = tk.Frame(self, bg=PRIMARY_COLOR, height=70)
        self.header_frame.pack(fill="x")
        self.header_frame.pack_propagate(False)
        
        header_font = Font(family="Inter Regular", size=16, weight="bold")
        tk.Label(
            self.header_frame, 
            text="pyTracker Model Selection", 
            font=header_font, 
            bg=PRIMARY_COLOR, 
            fg=TEXT_COLOR
        ).pack(pady=20)
        
        self.main_frame = tk.Frame(self, bg=BG_COLOR)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Instructions label
        tk.Label(
            self.main_frame,
            text="Select an LLM model to use with pyTracker:",
            font=("Inter Regular", 12),
            bg=BG_COLOR,
            fg=TEXT_COLOR
        ).pack(pady=(0, 10))
        
        # Frame for model cards
        models_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        models_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Define model families
        models = [
            {"name": "Qwen", "image": "public/logos/qwen-color.png", "id": "qwen"},
            {"name": "Llama", "image": "public/logos/llama-color.png", "id": "llama"},
            {"name": "Mistral", "image": "public/logos/mistral-color.png", "id": "mistral"},
            {"name": "Gemma", "image": "public/logos/gemma-color.png", "id": "gemma"},
            {"name": "Phi", "image": "public/logos/phi-color.png", "id": "phi"},
            {"name": "Llava", "image": "public/logos/llava-color.png", "id": "llava"}
        ]
        
        # Calculate columns - use 3 columns for better visibility
        cols = min(4, len(models))
        
        # Create a grid of model buttons
        for i, model in enumerate(models):
            # Calculate grid position
            row = i // cols
            col = i % cols
            
            # Create container frame with FIXED size
            container = tk.Frame(
                models_frame,
                bg=BG_COLOR,
                width=100,
                height=100,
                bd=0
            )
            container.grid(row=row*2, column=col, padx=15, pady=10)
            container.grid_propagate(False)
            
            # Create a canvas with rounded corners for the border
            canvas = tk.Canvas(
                container,
                bg=BG_COLOR,
                highlightthickness=0,
                width=100,
                height=100
            )
            canvas.place(relx=0.5, rely=0.5, anchor="center")
            
            # Draw rounded rectangle for border
            border_id = canvas.create_rounded_rectangle(
                10, 10, 90, 90,  # Coordinates
                radius=10,       # Radius for rounded corners
                fill=BG_COLOR,
                outline=PRIMARY_COLOR,
                width=2
            )
            
            # Load image
            try:
                img = Image.open(model["image"])
                img = img.resize((70, 70), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                # Create larger image for hover state
                hover_img = Image.open(model["image"])
                hover_img = hover_img.resize((80, 80), Image.LANCZOS)
                hover_photo = ImageTk.PhotoImage(hover_img)
                
                # Add image to canvas
                image_id = canvas.create_image(50, 50, image=photo, anchor="center")
                canvas.image = photo  # Keep reference
                canvas.hover_image = hover_photo  # Keep reference
                
                # Add hover effects without moving other elements
                def on_enter(e, canv=canvas, h_img=hover_photo, b_id=border_id):
                    canv.itemconfig(b_id, outline=ACCENT_COLOR, width=3)
                    canv.itemconfig(canv.find_withtag("current")[0], image=h_img)
                    
                def on_leave(e, canv=canvas, o_img=photo, b_id=border_id):
                    canv.itemconfig(b_id, outline=PRIMARY_COLOR, width=2)
                    canv.itemconfig(canv.find_withtag("current")[0], image=o_img)
                    
                canvas.tag_bind(image_id, "<Enter>", on_enter)
                canvas.tag_bind(image_id, "<Leave>", on_leave)
                canvas.tag_bind(image_id, "<Button-1>", 
                            lambda e, m=model["id"]: self.gotoModelVersionSelection(m))
                
                # Add label
                tk.Label(
                    models_frame,
                    text=model["name"],
                    fg=TEXT_COLOR,
                    bg=BG_COLOR,
                    font=("Inter Regular", 11)
                ).grid(row=row*2+1, column=col)
                
            except Exception as e:
                print(f"Error loading image for {model['name']}: {e}")
        
        # Add back button to return to previous screen
        back_button = tk.Button(
            self.main_frame,
            text="← Back",
            bg=BG_COLOR,
            fg=TEXT_COLOR,
            relief="flat",
            borderwidth=1,
            command=self.getModel
        )
        back_button.pack(side="bottom", pady=10)
        
        # Status message at the bottom
        self.status_var = tk.StringVar()
        self.status_var.set("Select a model to continue")
        
        self.status_label = ttk.Label(
            self, 
            textvariable=self.status_var,
            foreground=TEXT_COLOR,
            background=BG_COLOR
        )
        self.status_label.pack(side="bottom", pady=10)


def get_credentials():
    """
    Get Google API credentials, handling expired tokens gracefully.
    Automatically deletes invalid tokens and creates new ones when needed.
    
    Returns:
        Google API credentials object
    """
    try:
        with open(CONFIG_FILE, 'r') as f:
            config = json.load(f)
    except Exception as e:
        print(f"Error loading config: {e}")
        return None

    # The file token.json stores the user's access and refresh tokens
    token_path = os.path.join(CONFIG_DIR, 'token.json')
    creds = None
    token_valid = True
    
    # Check if token.json exists with valid credentials
    if os.path.exists(token_path):
        try:
            # Try to load existing credentials
            creds = Credentials.from_authorized_user_info(
                json.load(open(token_path)), SCOPES)
                
            # Test if credentials are valid by making a small request
            if creds.valid:
                try:
                    # Quick test to see if token actually works
                    service = build('sheets', 'v4', credentials=creds)
                    service.spreadsheets().get(spreadsheetId="dummy").execute()
                except Exception:
                    # If test request fails, mark token as invalid
                    print("Token appears to be invalid despite being marked as valid")
                    token_valid = False
                    
        except Exception as e:
            # If there's any error loading or using the token
            print(f"Error with existing token: {e}")
            token_valid = False
    
    # If there's no valid token, we need new credentials
    if not creds or not token_valid:
        # If token exists but is invalid, try to refresh it
        if creds and creds.expired and creds.refresh_token and token_valid:
            try:
                creds.refresh(Request())
                print("Successfully refreshed expired token")
            except Exception as e:
                print(f"Failed to refresh token: {e}")
                token_valid = False
                creds = None  # Clear the invalid credentials
        
        # If token couldn't be refreshed or wasn't valid to begin with, create a new one
        if not token_valid or not creds:
            # First remove the invalid token if it exists
            if os.path.exists(token_path):
                try:
                    os.remove(token_path)
                    print(f"Removed invalid token file: {token_path}")
                except Exception as e:
                    print(f"Error removing invalid token file: {e}")
            
            # Now get new credentials via OAuth flow
            try:
                credentials_path = os.path.join(CONFIG_DIR, 'credentials.json')
                if not os.path.exists(credentials_path):
                    print(f"Error: Credentials file not found at {credentials_path}")
                    return None
                    
                flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
                creds = flow.run_local_server(port=0)
                
                # Save the new credentials
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
                print("Generated and saved new token")
            except Exception as e:
                print(f"Error generating new token: {e}")
                return None
    
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
        applications_sheet.update('A1:H1', [['Status', 'Company', 'Date Applied', 'Last Updated', 'Link', 'Role', 'Company ID', 'Job ID']])
        applications_sheet.freeze(rows=1)
        applications_sheet.format('A1:H1', {'textFormat': {'bold': True}})
        applications_sheet.format('A1:H1', {"backgroundColor": {"red": 0.961, "green": 0.961, "blue": 0.961}})
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
    try:
        # Initialize fonts
        pyglet.font.add_file('public/fonts/Inter-VariableFont_opsz,wght.ttf')
    except Exception as e:
        print(f"Warning: Could not load font: {e}")
    
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


if __name__ == "__main__":
    main()