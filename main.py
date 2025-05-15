import tkinter as tk
from tkinter import simpledialog
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime
import os

def get_login_info():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    username = simpledialog.askstring("Login", "Enter your username:", parent=root)
    password = simpledialog.askstring("Login", "Enter your password:", show="*", parent=root)
    return username, password

def get_clean_text(element):
    # Get only the text content before any child elements
    text = element.inner_text().split('\n')[0].strip()
    return text

def get_download_path():
    # Get the user's desktop path
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    # Create a downloads folder if it doesn't exist
    download_path = os.path.join(desktop, "Decopress_Downloads")
    os.makedirs(download_path, exist_ok=True)
    return download_path

LOGIN_URL = "https://intranet.decopress.com"
DASHBOARD_URL = "https://intranet.decopress.com/JobStatusList/JobStatusList.aspx"

def scrape_orders(page):
    orders = []
    current_page = 1
    max_pages = 2  # We only want first two pages
    
    # Wait for the table to be present and visible
    page.wait_for_selector("table.data-results", state="visible", timeout=30000)
    print("Table found, starting to scrape...")
    
    while current_page <= max_pages:
        print(f"Processing page {current_page}")
        # Wait for any loading indicators to disappear
        page.wait_for_load_state('networkidle')
        
        rows = page.query_selector_all("table.data-results tbody tr")
        print(f"Found {len(rows)} rows on current page")
        
        for row in rows:
            try:
                days_element = row.query_selector("span.js-days-to-due-date")
                if not days_element:
                    print("Days element not found in row")
                    continue
                    
                days_text = days_element.inner_text().strip()
                print(f"Days text found: {days_text}")
                
                try:
                    days = int(days_text)
                except ValueError:
                    print(f"Could not convert days text to integer: {days_text}")
                    continue
                
                if days in (0, 1, 2):
                    job_number = get_clean_text(row.query_selector("td:nth-child(1)"))
                    # Only keep numeric job numbers
                    if not job_number.isdigit():
                        continue
                        
                    job_status = get_clean_text(row.query_selector("td:nth-child(4)"))
                    # Keep only text after hyphen if it exists
                    if " - " in job_status:
                        job_status = job_status.split(" - ")[1]
                    
                    order = {
                        "Job Number": job_number,
                        "Customer": get_clean_text(row.query_selector("td:nth-child(2)")),
                        "Description": get_clean_text(row.query_selector("td:nth-child(3)")),
                        "Job Status": job_status,
                        "Order #": get_clean_text(row.query_selector("td:nth-child(5)")),
                        "Date In": get_clean_text(row.query_selector("td:nth-child(6)")),
                        "Ship Date": get_clean_text(row.query_selector("td:nth-child(7)")),
                        "Days Remaining": days
                    }
                    orders.append(order)
                    print(f"Added order with {days} days remaining")
            except Exception as e:
                print(f"Error processing row: {str(e)}")
                continue
        
        # Move to next page if we haven't reached max pages
        if current_page < max_pages:
            try:
                # Find the next page link
                next_page = page.query_selector(f"ul.pagination li[data-lp='{current_page + 1}'] a.page-link")
                if next_page:
                    print(f"Clicking page {current_page + 1}")
                    next_page.click()
                    page.wait_for_timeout(2000)  # Wait for page load
                    current_page += 1
                else:
                    print("Next page link not found")
                    break
            except Exception as e:
                print(f"Error navigating to next page: {str(e)}")
                break
        else:
            print("Reached maximum page limit")
            break
            
    print(f"Total orders found: {len(orders)}")
    return orders

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        # Login
        page.goto(LOGIN_URL)
        page.wait_for_load_state('networkidle')  # Wait for page to fully load
        try:
            username, password = get_login_info()
            page.wait_for_selector("#txt_Username", timeout=60000)
            page.fill("#txt_Username", username)
            page.fill("#txt_Password", password)
            page.click("#btn_Login")
            page.wait_for_selector("#jobStatusListResults", timeout=10000)
        except Exception as e:
            print(f"Login failed: {str(e)}")
            browser.close()
            return

        # Go to Job Status List
        page.goto(DASHBOARD_URL)
        page.wait_for_selector("table.data-results")

        # Scrape
        print("Scraping urgent orders...")
        orders = scrape_orders(page)

        # Export to Excel
        if orders:
            df = pd.DataFrame(orders)
            df.sort_values("Days Remaining", inplace=True)
            
            # Create a title row
            title_df = pd.DataFrame([["DECOPRESS DAILY ORDERS"]], columns=["Title"])
            
            # Get current date for filename
            current_date = datetime.now().strftime("%Y-%m-%d")
            download_path = get_download_path()
            filename = f"{current_date}_DECOPRESS_DAILY_ORDERS.xlsx"
            filepath = os.path.join(download_path, filename)
            
            # Write to Excel with title
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                title_df.to_excel(writer, index=False, header=False)
                df.to_excel(writer, index=False, startrow=2)  # Start after title
                
            print(f"✅ Exported {len(orders)} urgent orders to Excel: {filepath}")
        else:
            print("⚠️ No 0, 1, or 2-day orders found.")

        browser.close()

if __name__ == "__main__":
    run()
