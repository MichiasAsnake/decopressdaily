import tkinter as tk
from tkinter import simpledialog, messagebox
import os
from datetime import datetime
import json
import base64

def _simple_encrypt(text):
    """Simple encoding to avoid storing plaintext passwords"""
    return base64.b64encode(text.encode()).decode()

def _simple_decrypt(encrypted_text):
    """Decode the encrypted text"""
    try:
        return base64.b64decode(encrypted_text.encode()).decode()
    except:
        return ""

def save_credentials(username, password):
    """Save credentials to a file with simple encryption"""
    credentials = {
        "username": username,
        "password": _simple_encrypt(password)
    }
    
    # Get the application data directory
    app_data_dir = os.path.join(os.path.expanduser("~"), ".decopress")
    os.makedirs(app_data_dir, exist_ok=True)
    
    # Path to credentials file
    credentials_file = os.path.join(app_data_dir, "credentials.json")
    
    # Save to file
    with open(credentials_file, "w") as f:
        json.dump(credentials, f)
    
    print("✅ Credentials saved")

def load_credentials():
    """Load saved credentials if they exist"""
    # Get the application data directory
    app_data_dir = os.path.join(os.path.expanduser("~"), ".decopress")
    credentials_file = os.path.join(app_data_dir, "credentials.json")
    
    if os.path.exists(credentials_file):
        try:
            with open(credentials_file, "r") as f:
                credentials = json.load(f)
                
            username = credentials.get("username", "")
            encrypted_password = credentials.get("password", "")
            
            if username and encrypted_password:
                password = _simple_decrypt(encrypted_password)
                return username, password
        except Exception as e:
            print(f"Error loading credentials: {str(e)}")
    
    return None, None

def get_login_info():
    """Get username and password from user, with option to save"""
    # Try to load saved credentials first
    saved_username, saved_password = load_credentials()
    
    # If we have saved credentials, ask if user wants to use them
    if saved_username and saved_password:
        print(f"Found saved credentials for user: {saved_username}")
        
        # Create a new root window for the dialog
        root = tk.Tk()
        root.withdraw()
        
        # Make sure this dialog appears on top
        root.attributes('-topmost', True)
        
        # Create a more noticeable dialog
        use_saved = messagebox.askyesno(
            "Saved Login Found", 
            f"Use saved login for user '{saved_username}'?\n\nClick 'No' to enter different credentials.",
            parent=root,
            icon=messagebox.QUESTION
        )
        
        print(f"User chose to use saved credentials: {use_saved}")
        
        if use_saved:
            root.destroy()
            return saved_username, saved_password
        
        # If user said no, continue to regular login
        root.destroy()
    else:
        print("No saved credentials found - showing login dialog")
    
    # Use simple dialogs instead of custom dialog to prevent potential stalling
    root = tk.Tk()
    root.withdraw()
    
    # Make sure dialog appears on top
    root.attributes('-topmost', True)
    
    # Use simple dialogs which are more reliable
    username = simpledialog.askstring("Login Required", "Enter your username:", parent=root)
    if not username:
        print("Login canceled - no username entered")
        root.destroy()
        return None, None
        
    password = simpledialog.askstring("Login Required", "Enter your password:", show="*", parent=root)
    if not password:
        print("Login canceled - no password entered")
        root.destroy()
        return None, None
    
    # Ask about saving credentials
    save_creds = messagebox.askyesno("Remember Login", 
                                     "Would you like to save your login credentials for next time?", 
                                     parent=root)
    if save_creds:
        save_credentials(username, password)
        print(f"Saved credentials for user: {username}")
    
    root.destroy()
    print(f"Returning credentials: Username={'*'*(len(username))}")
    
    return username, password

def get_clean_text(element):
    """Get only the text content before any child elements"""
    text = element.inner_text().split('\n')[0].strip()
    return text

def get_download_path():
    """Get path to save downloads"""
    # Get the user's desktop path
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    # Create a downloads folder if it doesn't exist
    download_path = os.path.join(desktop, "Decopress_Downloads")
    os.makedirs(download_path, exist_ok=True)
    return download_path

def get_current_date_formatted(format="%Y-%m-%d"):
    """Get current date in specified format"""
    return datetime.now().strftime(format)

def get_job_number():
    """Get job number from user"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    job_number = simpledialog.askstring("Job Number", "Enter the job number:", parent=root)
    if job_number and job_number.strip().isdigit():
        return job_number.strip()
    return None

def get_shipment_details(job_number):
    """Get shipment details from user"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Ask if this is a partial shipment
    is_partial = messagebox.askyesno("Partial Shipment", "Is this a partial shipment?", parent=root)
    
    shipment_details = {}
    
    if is_partial:
        # Get partial shipment details
        partial_of = simpledialog.askstring("Partial Shipment", "Enter partial number (e.g., '1 of 3'):", parent=root)
        if partial_of:
            shipment_details["partial_shipment"] = partial_of
    
    # Get order quantity
    order_qty = simpledialog.askstring("Order Quantity", "Enter the order quantity:", parent=root)
    if order_qty and order_qty.strip().isdigit():
        shipment_details["order_qty"] = order_qty.strip()
    
    # Get ship quantity
    ship_qty = simpledialog.askstring("Ship Quantity", "Enter the ship quantity:", parent=root)
    if ship_qty and ship_qty.strip().isdigit():
        shipment_details["ship_qty"] = ship_qty.strip()
    
    # Get number of boxes
    num_boxes = simpledialog.askstring("Number of Boxes", "Enter the number of boxes:", parent=root)
    if num_boxes and num_boxes.strip().isdigit():
        shipment_details["num_boxes"] = num_boxes.strip()
    
    # Get comments
    comments = simpledialog.askstring("Comments", "Enter any comments (optional):", parent=root)
    if comments:
        shipment_details["comments"] = comments.strip()
    
    return shipment_details

# URLs
LOGIN_URL = "https://intranet.decopress.com"
DASHBOARD_URL = "https://intranet.decopress.com/JobStatusList/JobStatusList.aspx"
JOB_URL_TEMPLATE = "https://intranet.decopress.com/Jobs/job.aspx?ID={}" 