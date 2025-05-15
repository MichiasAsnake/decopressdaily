"""
Fix Playwright browser handling in bundled applications.
This script will modify daily_orders.py and packing_slip.py to handle 
Playwright browser installation in a bundled application.
"""
import os
import shutil

def modify_daily_orders():
    """Modify daily_orders.py to handle Playwright in bundled app"""
    with open("daily_orders.py", "r") as f:
        content = f.read()
    
    # Add imports
    if "import sys" not in content:
        content = "import sys\nimport subprocess\nimport tempfile\n" + content
    
    # Add browser handling code
    browser_code = """
def ensure_browser_installed():
    """Ensure Playwright browsers are installed in bundled app"""
    if getattr(sys, 'frozen', False):
        print("Running in a bundled application - checking Playwright setup...")
        try:
            # Try importing playwright to see if it works normally
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                try:
                    browser = p.chromium.launch()
                    browser.close()
                    print("Playwright browser launched successfully!")
                    return
                except Exception as e:
                    if "Executable doesn't exist" not in str(e):
                        print(f"Error launching browser: {e}")
                        return
                    print("Playwright browser not installed, will install now...")
            
            # Need to install the browser
            temp_dir = tempfile.gettempdir()
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = temp_dir
            
            print(f"Installing Playwright browser to {temp_dir}")
            result = subprocess.run(
                ["playwright", "install", "chromium"],
                capture_output=True,
                text=True
            )
            if result.returncode != 0:
                print(f"Error installing browser: {result.stderr}")
            else:
                print("Playwright browser installed successfully!")
        except Exception as e:
            print(f"Error setting up Playwright: {e}")
    else:
        print("Running in development mode, using normal Playwright setup")
"""
    
    # Check if function is already added
    if "def ensure_browser_installed():" not in content:
        # Find position to insert (after imports)
        import_end = max([content.find("\n\n", content.find(imp)) for imp in ["import ", "from "]])
        if import_end > 0:
            content = content[:import_end+2] + browser_code + content[import_end+2:]
        else:
            content = browser_code + content

    # Add browser installation in run function
    run_function = "def run():"
    if run_function in content and "ensure_browser_installed()" not in content:
        run_pos = content.find(run_function)
        run_body_start = content.find(":", run_pos) + 1
        # Find the first line after the function def (with indentation)
        next_line_pos = content.find("\n", run_body_start) + 1
        # Get the indentation level
        indentation = ""
        for char in content[next_line_pos:]:
            if char.isspace():
                indentation += char
            else:
                break
        
        # Insert browser installation call
        browser_install_call = f"\n{indentation}# Ensure browser is installed\n{indentation}ensure_browser_installed()"
        content = content[:next_line_pos] + browser_install_call + content[next_line_pos:]
    
    # Modify browser launch to use temp directory
    launch_line = "browser = p.chromium.launch(headless=False)"
    new_launch_line = """browser = p.chromium.launch(
            headless=False,
            executable_path=None  # Let Playwright find the executable
        )"""
    
    if launch_line in content:
        content = content.replace(launch_line, new_launch_line)
    
    # Save the modified file
    with open("daily_orders.py", "w") as f:
        f.write(content)
    
    print("Modified daily_orders.py to handle Playwright in bundled app")

def modify_packing_slip():
    """Modify packing_slip.py to handle Playwright in bundled app"""
    with open("packing_slip.py", "r") as f:
        content = f.read()
    
    # Add imports
    if "import sys" not in content:
        content = "import sys\nimport subprocess\nimport tempfile\n" + content
    
    # Add browser handling code
    browser_code = """
def ensure_browser_installed():
    """Ensure Playwright browsers are installed in bundled app"""
    if getattr(sys, 'frozen', False):
        print("Running in a bundled application - checking Playwright setup...")
        try:
            # Try importing playwright to see if it works normally
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                try:
                    browser = p.chromium.launch()
                    browser.close()
                    print("Playwright browser launched successfully!")
                    return
                except Exception as e:
                    if "Executable doesn't exist" not in str(e):
                        print(f"Error launching browser: {e}")
                        return
                    print("Playwright browser not installed, will install now...")
            
            # Need to install the browser
            temp_dir = tempfile.gettempdir()
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = temp_dir
            
            print(f"Installing Playwright browser to {temp_dir}")
            result = subprocess.run(
                ["playwright", "install", "chromium"],
                capture_output=True,
                text=True
            )
            if result.returncode != 0:
                print(f"Error installing browser: {result.stderr}")
            else:
                print("Playwright browser installed successfully!")
        except Exception as e:
            print(f"Error setting up Playwright: {e}")
    else:
        print("Running in development mode, using normal Playwright setup")
"""
    
    # Check if function is already added
    if "def ensure_browser_installed():" not in content:
        # Find position to insert (after imports)
        import_end = max([content.find("\n\n", content.find(imp)) for imp in ["import ", "from "]])
        if import_end > 0:
            content = content[:import_end+2] + browser_code + content[import_end+2:]
        else:
            content = browser_code + content

    # Add browser installation in run function
    run_function = "def run():"
    if run_function in content and "ensure_browser_installed()" not in content:
        run_pos = content.find(run_function)
        run_body_start = content.find(":", run_pos) + 1
        # Find the first line after the function def (with indentation)
        next_line_pos = content.find("\n", run_body_start) + 1
        # Get the indentation level
        indentation = ""
        for char in content[next_line_pos:]:
            if char.isspace():
                indentation += char
            else:
                break
        
        # Insert browser installation call
        browser_install_call = f"\n{indentation}# Ensure browser is installed\n{indentation}ensure_browser_installed()"
        content = content[:next_line_pos] + browser_install_call + content[next_line_pos:]
    
    # Modify browser launch to use temp directory
    launch_line = "browser = p.chromium.launch(headless=False)"
    new_launch_line = """browser = p.chromium.launch(
            headless=False,
            executable_path=None  # Let Playwright find the executable
        )"""
    
    if launch_line in content:
        content = content.replace(launch_line, new_launch_line)
    
    # Save the modified file
    with open("packing_slip.py", "w") as f:
        f.write(content)
    
    print("Modified packing_slip.py to handle Playwright in bundled app")

if __name__ == "__main__":
    # Backup files first
    if not os.path.exists("backups"):
        os.makedirs("backups")
    
    # Backup daily_orders.py
    if os.path.exists("daily_orders.py"):
        shutil.copy2("daily_orders.py", "backups/daily_orders.py.bak")
    
    # Backup packing_slip.py
    if os.path.exists("packing_slip.py"):
        shutil.copy2("packing_slip.py", "backups/packing_slip.py.bak")
    
    # Modify files
    modify_daily_orders()
    modify_packing_slip()
    
    print("Files modified successfully. Backups saved in backups/ directory.") 