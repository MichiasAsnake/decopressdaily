"""
Script to find the Playwright browser paths for PyInstaller.
This will print the paths that need to be included in the PyInstaller spec file.
"""
import os
import sys
import json
from pathlib import Path

def find_playwright_browser_path():
    """Find the path to Playwright browser binaries."""
    try:
        # Try to locate playwright cache directory
        home_dir = os.path.expanduser("~")
        ms_playwright_path = os.path.join(home_dir, ".cache", "ms-playwright")
        
        # On Windows, it might be in appdata
        if sys.platform == "win32":
            ms_playwright_path = os.path.join(os.environ.get("APPDATA", ""), "ms-playwright")
        
        # Check if the directory exists
        if not os.path.exists(ms_playwright_path):
            print(f"Playwright browser directory not found at: {ms_playwright_path}")
            return None
        
        # Check for chromium directory
        chromium_path = os.path.join(ms_playwright_path, "chromium-")
        
        # Find chromium directories (may have version suffix)
        chromium_dirs = [d for d in os.listdir(ms_playwright_path) if d.startswith("chromium-")]
        
        if not chromium_dirs:
            print("Chromium directory not found in Playwright cache")
            return None
        
        # Get the first chromium directory
        chromium_path = os.path.join(ms_playwright_path, chromium_dirs[0])
        
        if not os.path.exists(chromium_path):
            print(f"Chromium path not found: {chromium_path}")
            return None
        
        print("\n=== Playwright Browser Information ===")
        print(f"Playwright cache directory: {ms_playwright_path}")
        print(f"Chromium directory: {chromium_path}")
        
        # Generate datas entries for PyInstaller
        print("\n=== Add these entries to your PyInstaller spec file ===")
        
        # For Windows, the main executable is in a different location
        if sys.platform == "win32":
            browser_path = os.path.join(chromium_path, "chrome-win")
            
            # Generate data entries
            data_entries = []
            
            # Add the main chromium directory
            data_entries.append(f"(r'{browser_path}', 'chromium')")
            
            # Add specific files needed
            exe_path = os.path.join(browser_path, "chrome.exe")
            if os.path.exists(exe_path):
                exe_rel_path = os.path.relpath(exe_path, os.getcwd())
                data_entries.append(f"(r'{exe_rel_path}', 'chromium')")
            
            print("datas=[")
            for entry in data_entries:
                print(f"    {entry},")
            print("]")
            
            return browser_path
            
        # For other platforms (Linux/Mac)
        else:
            # Different paths for different platforms
            if sys.platform == "darwin":  # macOS
                browser_path = os.path.join(chromium_path, "chrome-mac")
            else:  # Linux
                browser_path = os.path.join(chromium_path, "chrome-linux")
                
            print(f"(r'{browser_path}', 'chromium'),")
            return browser_path
            
    except Exception as e:
        print(f"Error finding Playwright browser path: {str(e)}")
        return None

if __name__ == "__main__":
    find_playwright_browser_path() 