"""Find Playwright browser installation path"""
from playwright.sync_api import sync_playwright
import os
import json
import sys

def find_browser_paths():
    with sync_playwright() as p:
        # Get executable path
        try:
            browser = p.chromium.launch()
            executable_path = browser.executable_path
            browser.close()
            
            print(f"Chromium executable path: {executable_path}")
            
            # Get the browser directory (parent of the executable)
            browser_dir = os.path.dirname(executable_path)
            parent_dir = os.path.dirname(browser_dir)
            
            print(f"Browser directory: {browser_dir}")
            print(f"Parent directory: {parent_dir}")
            
            # Print instructions for spec file
            print("\nAdd these to your PyInstaller spec file:")
            print("datas=[")
            print(f"    (r'{browser_dir}', 'chromium'),")
            print("]")
            
            return browser_dir
            
        except Exception as e:
            print(f"Error: {e}")
            return None

if __name__ == "__main__":
    find_browser_paths() 