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
        # Order Number
        order_number_elem = page.query_selector("input#orderNumber")
        if order_number_elem:
            shipping_info["Order #"] = order_number_elem.get_attribute("value") or ""
        
        # Description 
        desc_elem = page.query_selector("input#orderDescription")
        if desc_elem:
            shipping_info["Description"] = desc_elem.get_attribute("value") or ""
        
        # Customer name
        customer_name = page.query_selector("input#customer")
        if customer_name:
            shipping_info["Customer"] = customer_name.get_attribute("value") or ""
        
        # Customer contact - this should go in G13
        selected_contact_element = page.query_selector("select#customerUser option[selected]:not([hidden])")
        if selected_contact_element:
            shipping_info["Selected Contact"] = selected_contact_element.inner_text().strip()
        
        # Extract shipping info from the shipment-info list (for cell E6)
        shipment_info_items = []
        shipment_info_list = page.query_selector("ul.shipment-info")
        if shipment_info_list:
            # Extract company name
            company_elem = shipment_info_list.query_selector("li.media:first-child div.media-body")
            if company_elem:
                company_name = company_elem.inner_text().strip()
                shipment_info_items.append(company_name)
            
            # Extract address
            address_elem = shipment_info_list.query_selector("li.media address.mb-1")
            if address_elem:
                # Get just the text, not any nested elements (like the verified tag)
                address_text = address_elem.inner_text().split("Verified")[0].strip()
                shipment_info_items.append(address_text)
            
            # Extract contact/reference
            contact_elem = shipment_info_list.query_selector("li.media:nth-child(3) div.media-body")
            if contact_elem:
                contact_text = contact_elem.inner_text().strip()
                shipment_info_items.append(contact_text)
            
            # Extract shipping notes
            notes_elem = shipment_info_list.query_selector("li.media div.shipment-notes-container")
            if notes_elem:
                notes_text = notes_elem.inner_text().strip()
                shipment_info_items.append(notes_text)
            
            # Join all shipment info for cell E6
            shipping_info["Full Shipment Info"] = "\n".join(shipment_info_items)
            
            # Also keep individual parts for other potential uses
            if shipment_info_items:
                if len(shipment_info_items) > 0:
                    shipping_info["Ship To Company"] = shipment_info_items[0]
                if len(shipment_info_items) > 1:
                    shipping_info["Ship To Address"] = shipment_info_items[1]
                if len(shipment_info_items) > 2:
                    shipping_info["Ship To Reference"] = shipment_info_items[2]
                if len(shipment_info_items) > 3:
                    shipping_info["Ship To Notes"] = shipment_info_items[3]
        
        # Extract asset/SKU data from joblines
        assets = []
        # Look for all joblines that are not "GSORT" or other standard types
        jobline_rows = page.query_selector_all("table.job-joblines-list tr.js-jobline-row")
        for row in jobline_rows:
            # Get asset/SKU (in first column, but we want the asset tag, not "GSORT")
            asset_link = row.query_selector("td:first-child a.js-view-asset")
            
            if asset_link:  # Only process rows with asset links (not GSORT, etc)
                asset_tag = asset_link.inner_text().strip()
                
                # Get description (second column)
                description_cell = row.query_selector("td:nth-child(2)")
                description = description_cell.inner_text().strip() if description_cell else ""
                
                # Get quantity from the 4th column (quantity column)
                qty_cell = row.query_selector("td:nth-child(4)")
                qty = qty_cell.inner_text().strip() if qty_cell else ""
                
                # Only add if it's an asset SKU (has letters and numbers)
                if asset_tag and re.search(r'[A-Za-z]', asset_tag) and re.search(r'[0-9]', asset_tag):
                    print(f"Found asset: {asset_tag}, description: {description}, qty: {qty}")
                    assets.append({
                        "asset_tag": asset_tag,
                        "description": description,
                        "qty": qty
                    })
        
        shipping_info["assets"] = assets
        
        # Get job number from URL
        shipping_info["Job Number"] = job_number
        
    except Exception as e:
        print(f"Error extracting shipping information: {str(e)}")
    
    return shipping_info

def set_cell_value_safely(sheet, cell_reference, value):
    """
    Safely set a cell value, handling merged cells.
    For merged cells, it finds the top-left (primary) cell of the merged range and sets that instead.
    """
    cell = sheet[cell_reference]
    
    # Check if this is a merged cell
    for merged_range in sheet.merged_cells.ranges:
        if cell.coordinate in merged_range:
            # Get the top-left cell of the merged range (the primary cell)
            primary_cell_coords = merged_range.coord.split(':')[0]
            print(f"Cell {cell_reference} is part of merged range {merged_range.coord}, using {primary_cell_coords} instead")
            sheet[primary_cell_coords] = value
            return
    
    # Not a merged cell, set value directly
    cell.value = value

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
    
    # Ask for ship date from user
    root = tk.Tk()
    root.withdraw()
    ship_date = simpledialog.askstring("Ship Date", "Enter the ship date (MM/DD/YYYY):", parent=root)
    root.destroy()
    
    # Fill in the data using the safe method that handles merged cells
    # Date in G2 - NOT in A1
    current_date = get_current_date_formatted("%m/%d/%Y")
    set_cell_value_safely(sheet, "G2", current_date)
    
    # A13: Ship date (no text, just date)
    if ship_date:
        set_cell_value_safely(sheet, "A13", ship_date)  # Ship Date in A13
    
    # D13: ORDER# from input#orderNumber
    set_cell_value_safely(sheet, "D13", data.get("Order #", ""))  # Order # in D13
    
    # B13: JOB# 
    set_cell_value_safely(sheet, "B13", data.get("Job Number", ""))  # Job Number in B13
    
    # G13: Customer contact name (selected user)
    set_cell_value_safely(sheet, "G13", data.get("Selected Contact", ""))  # Customer contact in G13
    
    # Ship To info - Put everything in E6
    full_shipment_info = data.get("Full Shipment Info", "")
    if full_shipment_info:
        set_cell_value_safely(sheet, "E6", full_shipment_info)
    
    # Check for assets and add to the sheet
    assets = data.get("assets", [])
    if assets:
        # Get the order quantity from the first asset to show in the dialog
        expected_qty = assets[0].get("qty", "") if assets else ""
        
        # Ask for quantities with reference to the expected quantity
        root = tk.Tk()
        root.withdraw()
        
        # Show the expected quantity in the dialog
        order_qty_prompt = f"Enter the order quantity (expected: {expected_qty}):"
        order_qty = simpledialog.askstring("Order Quantity", order_qty_prompt, parent=root)
        
        ship_qty_prompt = f"Enter the ship quantity (expected: {expected_qty}):"
        ship_qty = simpledialog.askstring("Ship Quantity", ship_qty_prompt, parent=root)
        
        # Get number of boxes
        num_boxes = simpledialog.askstring("Number of Boxes", "Enter the number of boxes:", parent=root)
        root.destroy()
        
        # Process each asset (putting only unique ones in the sheet)
        processed_assets = set()
        row_index = 16  # Start at row 16 for assets
        
        for asset in assets:
            asset_tag = asset.get("asset_tag")
            if asset_tag and asset_tag not in processed_assets:
                # Put asset tag in column A
                set_cell_value_safely(sheet, f"A{row_index}", asset_tag)
                
                # Put description in column B
                set_cell_value_safely(sheet, f"B{row_index}", asset.get("description", ""))
                
                # Add to processed set to avoid duplicates
                processed_assets.add(asset_tag)
                row_index += 1
                
        # Set quantities for the first row (row 16)
        if order_qty and order_qty.strip():
            set_cell_value_safely(sheet, "F16", order_qty.strip())  # ORDER QTY in F16
            set_cell_value_safely(sheet, "F28", order_qty.strip())  # Total ORDER QTY in F28
            
        if ship_qty and ship_qty.strip():
            set_cell_value_safely(sheet, "H16", ship_qty.strip())  # SHIP QTY in H16
            set_cell_value_safely(sheet, "H28", ship_qty.strip())  # Total SHIP QTY in H28
            
        if num_boxes and num_boxes.strip():
            set_cell_value_safely(sheet, "I16", num_boxes.strip())  # # BOXES in I16
            set_cell_value_safely(sheet, "I28", num_boxes.strip())  # Total # BOXES in I28
    else:
        # If no assets found, use the regular shipment details
        set_cell_value_safely(sheet, "F16", data.get("order_qty", ""))  # ORDER QTY in F16
        set_cell_value_safely(sheet, "H16", data.get("ship_qty", ""))  # SHIP QTY in H16
        set_cell_value_safely(sheet, "I16", data.get("num_boxes", ""))  # # BOXES in I16
        
        # Totals row
        set_cell_value_safely(sheet, "F28", data.get("order_qty", ""))  # Total ORDER QTY in F28
        set_cell_value_safely(sheet, "H28", data.get("ship_qty", ""))  # Total SHIP QTY in H28
        set_cell_value_safely(sheet, "I28", data.get("num_boxes", ""))  # Total # BOXES in I28
    
    # Partial shipment info (if applicable)
    if "partial_shipment" in data:
        set_cell_value_safely(sheet, "G5", f"Partial Shipment: {data['partial_shipment']}")
    
    # Comments (if provided)
    if "comments" in data:
        set_cell_value_safely(sheet, "A25", "Comments:")
        set_cell_value_safely(sheet, "B25", data["comments"])
    
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
        # Use headless mode by default to avoid focus/modal issues
        launch_options = {
            "headless": True,  # Run in headless mode to avoid focus issues
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
            
    return excel_path, pdf_path if 'pdf_path' in locals() and pdf_path else None

if __name__ == "__main__":
    run() 