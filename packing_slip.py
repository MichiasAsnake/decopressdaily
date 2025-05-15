import tkinter as tk
from tkinter import simpledialog
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import os
import re
import time
import shutil
import sys
import tempfile
from openpyxl import load_workbook
from utils import (
    get_login_info, get_clean_text, get_download_path, get_job_number,
    get_current_date_formatted, LOGIN_URL, DASHBOARD_URL, JOB_URL_TEMPLATE,
    get_shipment_details
)

# Try to import win32com for PDF conversion (Windows only)
try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

def ensure_browser_installed():
    """Ensure we can use a browser in bundled app"""
    if getattr(sys, 'frozen', False):
        print("Running in a bundled application - checking browser setup...")
        
        # Set a temp directory for Playwright browsers if needed
        temp_dir = tempfile.mkdtemp(prefix="decopress_browser_")
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = temp_dir
        
        # Look for Chrome/Edge installations on the system
        chrome_paths = [
            # Regular Chrome paths
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            # Microsoft Edge paths
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]
        
        # Check which browsers are available
        available_browsers = []
        for path in chrome_paths:
            if os.path.exists(path):
                available_browsers.append(path)
                print(f"Found browser: {path}")
        
        if available_browsers:
            os.environ["PLAYWRIGHT_CHROME_EXECUTABLE_PATH"] = available_browsers[0]
            print(f"Using browser: {available_browsers[0]}")
            return available_browsers[0]
        else:
            print("No system browsers found - will try using Playwright's default approach")
            return None
    else:
        print("Running in development mode, using normal Playwright setup")
        return None

def find_job_in_job_list(page, job_number):
    """Find job information in job status list."""
    max_pages = 10  # Maximum number of pages to search
    current_page = 1
    
    print(f"Searching for job number {job_number}...")
    job_info = {}
    
    while current_page <= max_pages:
        print(f"Searching on page {current_page}")
        # Wait for table to load
        page.wait_for_selector("table.data-results", state="visible", timeout=30000)
        page.wait_for_load_state('networkidle')
        
        # Find all rows
        rows = page.query_selector_all("table.data-results tbody tr")
        
        # Search for job number in each row
        for row in rows:
            try:
                row_job_number = get_clean_text(row.query_selector("td:nth-child(1)"))
                if row_job_number == job_number:
                    print(f"Found job {job_number} on page {current_page}")
                    job_info = {
                        "Job Number": job_number,
                        "Customer": get_clean_text(row.query_selector("td:nth-child(2)")),
                        "Description": get_clean_text(row.query_selector("td:nth-child(3)")),
                        "Job Status": get_clean_text(row.query_selector("td:nth-child(4)")),
                        "Order #": get_clean_text(row.query_selector("td:nth-child(5)")),
                        "Date In": get_clean_text(row.query_selector("td:nth-child(6)")),
                        "Ship Date": get_clean_text(row.query_selector("td:nth-child(7)")),
                    }
                    return job_info
            except Exception as e:
                print(f"Error processing row: {str(e)}")
                continue
        
        # Check if there's a next page
        next_page = page.query_selector(f"ul.pagination li[data-lp='{current_page + 1}'] a.page-link")
        if next_page:
            print(f"Moving to page {current_page + 1}")
            next_page.click()
            page.wait_for_timeout(2000)  # Wait for page load
            current_page += 1
        else:
            print("No more pages to search")
            break
    
    print(f"Job {job_number} not found after searching {current_page} pages")
    return None

def get_job_details(page, job_number):
    """Get detailed job information from the Job page."""
    job_url = JOB_URL_TEMPLATE.format(job_number)
    print(f"Navigating to job page: {job_url}")
    
    page.goto(job_url)
    page.wait_for_load_state('networkidle')
    
    # Extract shipping information
    shipping_info = {}
    try:
        # Customer name
        customer_name = page.query_selector("ul.shipment-info li.media:first-child div.media-body")
        if customer_name:
            shipping_info["Customer Name"] = customer_name.inner_text().strip()
        
        # Address
        address_element = page.query_selector("ul.shipment-info li.media address.mb-1")
        if address_element:
            shipping_info["Address"] = address_element.inner_text().strip()
        
        # Contact
        contact_element = page.query_selector("ul.shipment-info li.media:last-child div.media-body")
        if contact_element:
            contact_text = contact_element.inner_text().strip()
            # Try to extract name, phone and email
            name_match = re.search(r'^([^,]+)', contact_text)
            phone_match = re.search(r'\((\d{3})\)\s*(\d{3})-(\d{4})', contact_text)
            email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', contact_text)
            
            if name_match:
                shipping_info["Contact Name"] = name_match.group(1).strip()
            if phone_match:
                shipping_info["Phone"] = phone_match.group(0)
            if email_match:
                shipping_info["Email"] = email_match.group(1)
        
        # Selected contact
        selected_contact_element = page.query_selector("select#customerUser option[selected]:not([hidden])")
        if selected_contact_element:
            shipping_info["Selected Contact"] = selected_contact_element.inner_text().strip()
    except Exception as e:
        print(f"Error extracting shipping information: {str(e)}")
    
    return shipping_info

def create_packing_slip(job_info, shipping_info, shipment_details):
    """Create a packing slip Excel file using the template."""
    # Combine all information
    data = {**job_info, **shipping_info, **shipment_details}
    
    # Get the template path
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, "PackingSlipTemplate.xlsx")
    
    if not os.path.exists(template_path):
        print(f"❌ Template file not found: {template_path}")
        return None
    
    # Get the download path
    download_path = get_download_path()
    job_number = data.get("Job Number", "unknown")
    
    # Use job number as filename
    excel_filename = f"{job_number}.xlsx"
    pdf_filename = f"{job_number}.pdf"
    
    excel_filepath = os.path.join(download_path, excel_filename)
    pdf_filepath = os.path.join(download_path, pdf_filename)
    
    # Copy the template to the new file
    shutil.copy2(template_path, excel_filepath)
    
    # Open the template with openpyxl
    workbook = load_workbook(excel_filepath)
    sheet = workbook.active
    
    # Format date strings
    current_date = get_current_date_formatted("%m/%d/%Y")
    
    # Fill in the data
    # These cell references should match your template
    # Adjust these references based on the actual template
    
    # Header info
    sheet["G2"] = current_date  # Date
    sheet["A12"] = data.get("Ship Date", "")  # Ship Date
    sheet["B12"] = data.get("Order #", "")  # Order #
    sheet["C12"] = data.get("Description", "")  # PO #
    sheet["D12"] = data.get("Customer Name", data.get("Customer", ""))  # Customer
    
    # Ship To info
    sheet["D6"] = data.get("Customer Name", data.get("Customer", ""))  # Customer name
    
    # Address (combine lines)
    address = data.get("Address", "")
    address_lines = [line.strip() for line in address.split(",")]
    # Join all address lines into one cell
    sheet["D7"] = ", ".join(address_lines) if address_lines else ""
    
    # Contact info
    sheet["D12"] = data.get("Selected Contact", data.get("Contact Name", ""))  # Contact
    # Combine contact info on one line
    contact_name = data.get("Selected Contact", data.get("Contact Name", ""))
    phone = data.get("Phone", data.get("Contact Info", "").split(",")[1].strip() if data.get("Contact Info") else "")
    email = data.get("Email", "")
    sheet["D8"] = f"{contact_name}, {phone}, {email}"  # Combined contact info
    
    # Item info (first row)
    sheet["A18"] = data.get("Job Number", "")  # Item
    sheet["B18"] = data.get("Description", "")  # Description
    
    # Shipment details
    sheet["E18"] = data.get("order_qty", "")  # ORDER QTY
    sheet["G18"] = data.get("ship_qty", "")  # SHIP QTY
    sheet["H18"] = data.get("num_boxes", "")  # # BOXES
    
    # Copy values to row 29
    sheet["E30"] = data.get("order_qty", "")  # ORDER QTY
    sheet["G30"] = data.get("ship_qty", "")  # SHIP QTY 
    sheet["H30"] = data.get("num_boxes", "")  # # BOXES
    # Partial shipment info (if applicable)
    if "partial_shipment" in data:
        sheet["G5"] = f"Partial Shipment: {data['partial_shipment']}"
    
    # Comments (if provided)
    if "comments" in data:
        sheet["A25"] = "Comments:"
        sheet["B25"] = data["comments"]
    
    # Save the workbook
    workbook.save(excel_filepath)
    
    # Convert to PDF if possible
    pdf_created = False
    if HAS_WIN32COM:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(excel_filepath))
            ws = wb.Worksheets[0]
            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_filepath))
            wb.Close()
            excel.Quit()
            pdf_created = True
            print(f"✅ Created PDF: {pdf_filepath}")
        except Exception as e:
            print(f"❌ Error creating PDF: {str(e)}")
    
    print(f"✅ Created packing slip from template: {excel_filepath}")
    return excel_filepath, pdf_filepath if pdf_created else None

def run():
    """Main function to run the packing slip generation process."""
    # Get job number
    job_number = get_job_number()
    if not job_number:
        print("❌ Invalid job number provided")
        return
    
    # Get shipment details
    shipment_details = get_shipment_details(job_number)
    
    # Ensure browser is installed
    browser_path = ensure_browser_installed()
    
    with sync_playwright() as p:
        # Launch browser using system installation if available
        launch_options = {
            "headless": False,
        }
        
        if browser_path:
            launch_options["executable_path"] = browser_path
            browser = p.chromium.launch(**launch_options)
        else:
            # Let Playwright try to find its own browser
            browser = p.chromium.launch(**launch_options)
            
        context = browser.new_context()
        page = context.new_page()
        
        try:
            # Login
            page.goto(LOGIN_URL)
            page.wait_for_load_state('networkidle')
            
            username, password = get_login_info()
            
            # Check if login was cancelled
            if not username or not password:
                print("Login cancelled by user")
                browser.close()
                return
                
            page.wait_for_selector("#txt_Username", timeout=60000)
            page.fill("#txt_Username", username)
            page.fill("#txt_Password", password)
            page.click("#btn_Login")
            page.wait_for_selector("#jobStatusListResults", timeout=10000)
            
            # Go to Job Status List and find the job
            page.goto(DASHBOARD_URL)
            page.wait_for_selector("table.data-results")
            
            # Search for job in the job list
            job_info = find_job_in_job_list(page, job_number)
            if not job_info:
                print(f"❌ Job {job_number} not found")
                browser.close()
                return
            
            # Get detailed job information
            shipping_info = get_job_details(page, job_number)
            
            # Create packing slip
            excel_path, pdf_path = create_packing_slip(job_info, shipping_info, shipment_details)
            
            print(f"✅ Successfully created packing slip for job {job_number}")
            
            # Show success message with file paths
            results = f"Excel file: {excel_path}"
            if pdf_path:
                results += f"\nPDF file: {pdf_path}"
            else:
                results += "\nPDF export not available. Install pywin32 for PDF support."
            
            print(results)
            
        except Exception as e:
            print(f"❌ Error: {str(e)}")
        finally:
            browser.close()

if __name__ == "__main__":
    run() 