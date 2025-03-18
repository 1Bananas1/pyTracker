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
import gspread,pyglet
from tkinter import PhotoImage

# Add this near the top of your file with other imports
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
        
        self.manualModel()

        # Check if configuration exists and load appropriate screen
        # if os.path.exists(CONFIG_FILE):
        #     self.load_main_screen()
        # else:
        #     self.load_setup_screen()

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
        #print(totalFree)
        gpuList = pd.read_csv('public/data/ollama_nlp_models.csv')
        recommendedFiltered = gpuList.loc[gpuList['recVRAM'] <= totalFree]
        minFiltered = gpuList.loc[gpuList['VRAM'] <= totalFree]
        #print(recommendedFiltered,minFiltered)
        
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
                
                return selected
            except Exception as e:
                messagebox.showerror("Error", f"Could not save model selection: {str(e)}")
                return None
        else:
            messagebox.showwarning("No Models Found", "No compatible models were found for your GPU.")
            print("No compatible models found")
            return None
        
    def manualModel(self):
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
        
        # Create a canvas with scrollbar for the models
        canvas_frame = tk.Frame(self.main_frame, bg=BG_COLOR)
        canvas_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Add scrollbar
        scrollbar = tk.Scrollbar(canvas_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        
        # Create canvas
        canvas = tk.Canvas(
            canvas_frame, 
            bg=BG_COLOR,
            yscrollcommand=scrollbar.set,
            highlightthickness=0  # Remove border
        )
        canvas.pack(side="left", fill="both", expand=True)
        
        # Configure scrollbar to work with canvas
        scrollbar.config(command=canvas.yview)
        
        # Frame inside canvas to hold models
        models_frame = tk.Frame(canvas, bg=BG_COLOR)
        
        # Create a window within the canvas to contain the models frame
        canvas_window = canvas.create_window((0, 0), window=models_frame, anchor="nw")
        
        # Define your models
        models = [
            {"name": "Qwen", "image": "public/logos/qwen-color.png", "id": "qwen"},
            {"name": "Llama", "image": "public/logos/llama-color.png", "id": "llama"},
            {"name": "Mistral", "image": "public/logos/mistral-color.png", "id": "mistral"},
            {"name": "Gemma", "image": "public/logos/gemma-color.png", "id": "gemma"},
            {"name": "Phi", "image": "public/logos/phi-color.png", "id": "phi"}
            # Add more models as needed
        ]
        
        # Calculate columns dynamically (max 3 columns)
        cols = min(3, len(models))
        
        # Create a grid of model buttons
        for i, model in enumerate(models):
            # Calculate grid position
            row = i // cols
            col = i % cols
            
            # Create container frame with FIXED size to prevent layout shifts
            container = tk.Frame(
                models_frame,
                bg=BG_COLOR,  # Change to background color so it blends in
                width=100,    # Fixed width
                height=100,   # Fixed height
                bd=0
            )
            container.grid(row=row*2, column=col, padx=15, pady=10)
            container.grid_propagate(False)  # Prevent size changes
            
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
                            lambda e, m=model["id"]: self.select_model(m))
                
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
        
        # Ensure mousewheel scrolling works
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # Update scrollregion when all widgets are in place
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Set the width of the canvas window to the width of the canvas
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
        
        models_frame.bind("<Configure>", configure_scroll_region)
        
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
            command=self.automatic_setup,
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

    # Add this method to handle automatic setup
    def automatic_setup(self):
        # This is a placeholder - implement automatic setup logic here
        messagebox.showinfo("Automatic Setup", "Automatic setup not yet implemented.")
        # Potentially call a different setup function or skip certain steps

    def select_model(self, model_id):
        # Update configuration with selected model
        pass
        # try:
        #     config_data = {}
        #     if os.path.exists(CONFIG_FILE):
        #         with open(CONFIG_FILE, 'r') as f:
        #             config_data = json.load(f)
            
        #     config_data['model_id'] = model_id
            
        #     # Ensure config directory exists
        #     if not os.path.exists(CONFIG_DIR):
        #         os.makedirs(CONFIG_DIR)
            
        #     # Save config
        #     with open(CONFIG_FILE, 'w') as f:
        #         json.dump(config_data, f)
                
        #     self.status_var.set(f"Selected model: {model_id}")
        #     messagebox.showinfo("Model Selected", f"Model {model_id} has been selected")
            
        # except Exception as e:
        #     messagebox.showerror("Error", f"Could not save model selection: {str(e)}")

    def create_rounded_rectangle(self, canvas, x1, y1, x2, y2, radius=25, **kwargs):
        """Draw a rounded rectangle on a canvas."""
        points = [
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
            x1, y1
        ]
        return canvas.create_polygon(points, **kwargs, smooth=True)

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