import tkinter as tk
from tkinter import ttk, messagebox
import os
import daily_orders
import packing_slip
import random
import threading
import subprocess
from datetime import datetime
from PIL import Image, ImageTk
import json

class ModernTheme:
    """Modern theme colors and styles"""
    BG_COLOR = "#f5f5f7"
    PRIMARY_COLOR = "#0071e3"
    SECONDARY_COLOR = "#1d1d1f"
    ACCENT_COLOR = "#06c"
    
    BUTTON_STYLE = {
        "background": PRIMARY_COLOR,
        "foreground": "white",
        "activebackground": "#005bbf",
        "activeforeground": "white",
        "font": ("Segoe UI", 11),
        "borderwidth": 0,
        "relief": tk.FLAT,
        "padx": 15,
        "pady": 8,
        "cursor": "hand2"
    }
    
    TITLE_FONT = ("Segoe UI", 18, "bold")
    SUBTITLE_FONT = ("Segoe UI", 12)
    BUTTON_FONT = ("Segoe UI", 11)
    STATUS_FONT = ("Segoe UI", 9)

class LoadingScreen:
    """Loading screen with funny facts"""
    
    FUNNY_FACTS = [
        "Did you know? A patch is just a piece of fabric until it becomes awesome.",
        "Fun fact: The average person spends 2 weeks of their life waiting for Excel to load.",
        "Did you know? The first embroidered patch was created by accident when someone sneezed while sewing.",
        "Fun fact: If all the patches we've made were laid end to end, they would reach... well, not very far actually.",
        "Did you know? The most common password is '123456'. Please don't use that for your login.",
        "Fun fact: Smiling while waiting makes time go 42% faster. Science!",
        "Did you know? Patches were originally invented to cover holes, now they're cooler than the original garment.",
        "Fun fact: The average person blinks 15-20 times per minute. Unless they're staring at this loading screen.",
        "Did you know? Our software runs on coffee and determination.",
        "Fun fact: If you read this, you're officially part of the 'patient people' club.",
        "Did you know? The first computer bug was an actual bug - a moth trapped in a relay.",
        "Fun fact: You've now read more funny facts than most people do in a lifetime.",
        "Did you know? This loading screen was created just to entertain you. You're welcome!",
        "Fun fact: Clicking the screen repeatedly does not make it load faster. But it feels good.",
        "Did you know? The person who created this loading screen deserves a raise.",
    ]
    
    def __init__(self, parent):
        self.parent = parent
        
        # First withdraw the parent to prevent flashing
        parent.withdraw()
        
        # Create an independent toplevel window with higher visibility
        self.top = tk.Toplevel()  # Create without parent to make it more independent
        self.top.title("⚡ PROCESSING ORDERS ⚡")
        self.top.geometry("550x350")
        self.top.resizable(False, False)
        
        # Configure window appearance for maximum visibility
        self.top.configure(bg="#FFD700")  # Bright gold background
        
        # Make it stay on top but don't make it modal
        self.top.attributes('-topmost', False)  # Change to False to allow other windows to appear on top
        self.top.lift()
        
        # Prevent the window from being minimized or closed
        self.top.protocol("WM_DELETE_WINDOW", lambda: None)  # Disable close button
        self.top.attributes("-toolwindow", True)  # Make it a tool window (no minimize/maximize)
        
        # Add a thick border to make it more noticeable
        self.top.configure(highlightbackground="red", 
                          highlightcolor="red", 
                          highlightthickness=5)
        
        # Make the window transient but not modal
        self.top.transient(parent)
        # Explicitly do NOT use grab_set to allow other dialogs to function
        self.has_grab = False
        
        # Center the window on screen
        self.center_window()
        
        # Create loading UI
        self.setup_ui()
        
        # Start animation
        self.animate()
        
        # Start fact rotation
        self.rotate_facts()
        
        # Start border flashing
        self.flash_border()
        
        # Force visibility without grabbing focus
        for _ in range(5):  # Fewer update cycles to avoid potential issues
            try:
                self.top.update()
                # Don't try to grab focus or set topmost - allow other dialogs to appear
            except Exception as e:
                print(f"Error during update cycle: {e}")
                pass
    
    def center_window(self):
        # Get screen dimensions
        screen_width = self.top.winfo_screenwidth()
        screen_height = self.top.winfo_screenheight()
        
        # Calculate position
        width = 550
        height = 350
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # Set window position to center of screen
        self.top.geometry(f"{width}x{height}+{x}+{y}")
    
    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.top, bg="#FFD700", padx=25, pady=25)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title with attention-grabbing characters
        title_label = tk.Label(main_frame, 
                            text="⚡ PROCESSING ORDERS ⚡", 
                            font=("Arial Black", 18, "bold"),
                            bg="#FFD700",
                            fg="#FF0000")  # Red text
        title_label.pack(pady=(0, 20))
        
        # Smiley face (text-based)
        self.smiley_label = tk.Label(main_frame, 
                                 text=":)", 
                                 font=("Arial Black", 50),
                                 bg="#FFD700",
                                 fg="#1E90FF")  # Blue text
        self.smiley_label.pack(pady=(0, 20))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode="indeterminate", length=450)
        self.progress.pack(pady=(0, 20))
        
        # Funny fact label
        self.fact_label = tk.Label(main_frame, 
                               text=random.choice(self.FUNNY_FACTS),
                               font=("Arial", 12, "bold"),
                               wraplength=450,
                               justify="center",
                               bg="#FFD700",
                               fg="#000080")  # Navy blue text
        self.fact_label.pack(pady=(0, 10))
        
        # "Please wait" label
        please_wait = tk.Label(main_frame,
                           text="Please wait - do not close this window",
                           font=("Arial", 10, "italic"),
                           bg="#FFD700",
                           fg="#8B0000")  # Dark red
        please_wait.pack(pady=(10, 0))
        
        # Start progress bar immediately
        self.progress.start(8)  # Faster animation
    
    def flash_border(self):
        """Flash the border color to make the window more noticeable"""
        colors = ["red", "blue", "green", "purple", "orange"]
        current_color = 0
        
        def change_color():
            nonlocal current_color
            try:
                if self.top.winfo_exists():
                    current_color = (current_color + 1) % len(colors)
                    self.top.configure(highlightbackground=colors[current_color], 
                                      highlightcolor=colors[current_color])
                    self.top.update_idletasks()
                    # Use faster refresh for more attention-grabbing effect
                    self.top.after(300, change_color)
            except:
                # Window might have been destroyed
                pass
        
        # Start the color changing
        change_color()
    
    def animate(self):
        """Animate the progress bar and smiley face"""
        self.progress.start(8)  # Faster progress bar animation
        
        # More expressive smileys
        smileys = [":)", ";)", ":D", "^_^", ":P", "😊", "😎", "🤩", "😄", "👍"]
        self.current_smiley = 0
        
        def update_smiley():
            try:
                if self.top.winfo_exists():
                    self.current_smiley = (self.current_smiley + 1) % len(smileys)
                    self.smiley_label.config(text=smileys[self.current_smiley])
                    # Faster animation
                    self.top.after(800, update_smiley)
            except:
                pass  # Window might have been destroyed
        
        update_smiley()
    
    def rotate_facts(self):
        """Rotate through funny facts"""
        def update_fact():
            try:
                if self.top.winfo_exists():
                    current_fact = self.fact_label.cget("text")
                    new_fact = random.choice([f for f in self.FUNNY_FACTS if f != current_fact])
                    self.fact_label.config(text=new_fact)
                    self.top.after(4000, update_fact)  # Change fact every 4 seconds
            except:
                pass  # Window might have been destroyed
        
        update_fact()
    
    def close(self):
        """Close the loading screen"""
        try:
            # Stop animations
            if hasattr(self, 'progress') and self.progress.winfo_exists():
                self.progress.stop()
            
            # Now destroy the window - no need to release grab since we're not using it
            if hasattr(self, 'top') and self.top.winfo_exists():
                self.top.destroy()
                
            # Make sure parent is shown
            if hasattr(self, 'parent') and self.parent.winfo_exists():
                try:
                    self.parent.deiconify()
                    self.parent.update()
                except Exception as e:
                    print(f"Error showing parent window: {e}")
                    pass
        except Exception as e:
            print(f"Error closing loading screen: {e}")
            # If anything fails, try a more direct approach
            try:
                if hasattr(self, 'top') and self.top.winfo_exists():
                    self.top.destroy()
            except:
                pass

class DecoPressApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DecoPress Daily Tool")
        self.root.geometry("700x600")  # Increased height for recent files
        self.root.resizable(False, False)
        self.root.configure(bg=ModernTheme.BG_COLOR)
        
        # Recent files tracking
        self.recent_files = []
        self.load_recent_files()
        
        # Set up ttk style
        self.style = ttk.Style()
        self.style.theme_use('clam')  # Use a modern looking theme
        
        # Configure ttk styles
        self.style.configure('TFrame', background=ModernTheme.BG_COLOR)
        self.style.configure('Title.TLabel', 
                           font=ModernTheme.TITLE_FONT, 
                           background=ModernTheme.BG_COLOR, 
                           foreground=ModernTheme.SECONDARY_COLOR)
        self.style.configure('Subtitle.TLabel', 
                           font=ModernTheme.SUBTITLE_FONT, 
                           background=ModernTheme.BG_COLOR, 
                           foreground=ModernTheme.SECONDARY_COLOR)
        self.style.configure('Status.TLabel', 
                           font=ModernTheme.STATUS_FONT, 
                           background=ModernTheme.BG_COLOR, 
                           foreground="#666666")
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main container
        main_container = ttk.Frame(self.root, padding=30)
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Header frame
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=tk.X, pady=(0, 30))
        
        # Logo and title in the same row
        try:
            # Load and display logo if exists
            script_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(script_dir, "DecoPressLogo.jpg")
            
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((80, 80), Image.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                
                logo_label = ttk.Label(header_frame, image=logo_photo, background=ModernTheme.BG_COLOR)
                logo_label.image = logo_photo  # Keep a reference
                logo_label.pack(side=tk.LEFT, padx=(0, 20))
                
                # Title and subtitle
                title_box = ttk.Frame(header_frame)
                title_box.pack(side=tk.LEFT, fill=tk.Y, padx=0)
                
                title_label = ttk.Label(title_box, 
                                     text="DecoPress Daily Tool", 
                                     style='Title.TLabel')
                title_label.pack(anchor=tk.W)
                
                subtitle_label = ttk.Label(title_box, 
                                      text="Automated order tracking and packing slip generation", 
                                      style='Subtitle.TLabel')
                subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        except Exception as e:
            print(f"Could not load logo: {str(e)}")
            # Fallback to just title if logo fails
            title_label = ttk.Label(header_frame, 
                                 text="DecoPress Daily Tool", 
                                 style='Title.TLabel')
            title_label.pack(anchor=tk.CENTER, pady=(0, 5))
        
        # Separator
        separator = ttk.Separator(main_container, orient='horizontal')
        separator.pack(fill=tk.X, pady=(0, 30))
        
        # Content area
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create custom button style
        self.create_action_button = lambda parent, text, command: self.create_button(
            parent, text, command, is_primary=True
        )
        
        # Action cards
        # Daily Orders card
        daily_frame = ttk.Frame(content_frame, padding=15)
        daily_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        daily_title = ttk.Label(daily_frame, text="Daily Order List", 
                             font=("Segoe UI", 14, "bold"),
                             background=ModernTheme.BG_COLOR)
        daily_title.pack(anchor=tk.W, pady=(0, 10))
        
        daily_desc = ttk.Label(daily_frame, 
                           text="Generate a list of urgent orders due in the next 0-4 days. Keeps track of job status, customer details, and shipping dates.",
                           wraplength=280,
                           background=ModernTheme.BG_COLOR)
        daily_desc.pack(anchor=tk.W, pady=(0, 20), fill=tk.X)
        
        daily_btn = self.create_action_button(
            daily_frame, "Generate Daily Orders", self.run_daily_orders
        )
        daily_btn.pack(pady=10, fill=tk.X)
        
        # Vertical separator
        vsep = ttk.Separator(content_frame, orient='vertical')
        vsep.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # Packing Slip card
        packing_frame = ttk.Frame(content_frame, padding=15)
        packing_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        packing_title = ttk.Label(packing_frame, text="Packing Slip", 
                               font=("Segoe UI", 14, "bold"),
                               background=ModernTheme.BG_COLOR)
        packing_title.pack(anchor=tk.W, pady=(0, 10))
        
        packing_desc = ttk.Label(packing_frame, 
                             text="Create customized packing slips for any job. Includes shipping details, product information, and allows for partial shipment tracking.",
                             wraplength=280,
                             background=ModernTheme.BG_COLOR)
        packing_desc.pack(anchor=tk.W, pady=(0, 20), fill=tk.X)
        
        packing_btn = self.create_action_button(
            packing_frame, "Create Packing Slip", self.run_packing_slip
        )
        packing_btn.pack(pady=10, fill=tk.X)
        
        # Add a clear credentials button
        credentials_btn = self.create_button(
            main_container, "Clear Saved Credentials", self.clear_credentials, is_primary=False
        )
        credentials_btn.pack(pady=10, anchor=tk.CENTER)
        
        # Recent Files Section
        recent_files_frame = ttk.Frame(main_container)
        recent_files_frame.pack(fill=tk.X, pady=(20, 0))
        
        recent_title = ttk.Label(recent_files_frame, text="Recent Files", 
                              font=("Segoe UI", 12, "bold"),
                              background=ModernTheme.BG_COLOR)
        recent_title.pack(anchor=tk.W, pady=(0, 10))
        
        # Create a frame for the file list
        self.files_frame = ttk.Frame(recent_files_frame)
        self.files_frame.pack(fill=tk.X)
        
        # Populate recent files
        self.update_recent_files_ui()
        
        # Footer with status bar
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=10)
        
        # Version info
        version_label = ttk.Label(footer_frame, text="v1.0.0", style='Status.TLabel')
        version_label.pack(side=tk.LEFT, padx=20)
        
        # Date
        date_label = ttk.Label(footer_frame, 
                            text=f"Date: {datetime.now().strftime('%Y-%m-%d')}", 
                            style='Status.TLabel')
        date_label.pack(side=tk.RIGHT, padx=20)
    
    def load_recent_files(self):
        """Load recent files from JSON"""
        try:
            app_data_dir = os.path.join(os.path.expanduser("~"), ".decopress")
            recent_files_path = os.path.join(app_data_dir, "recent_files.json")
            
            if os.path.exists(recent_files_path):
                with open(recent_files_path, "r") as f:
                    self.recent_files = json.load(f)
                    # Keep only the 5 most recent files
                    self.recent_files = self.recent_files[:5]
        except Exception as e:
            print(f"Error loading recent files: {str(e)}")
            self.recent_files = []
    
    def save_recent_files(self):
        """Save recent files to JSON"""
        try:
            app_data_dir = os.path.join(os.path.expanduser("~"), ".decopress")
            os.makedirs(app_data_dir, exist_ok=True)
            recent_files_path = os.path.join(app_data_dir, "recent_files.json")
            
            with open(recent_files_path, "w") as f:
                json.dump(self.recent_files, f)
        except Exception as e:
            print(f"Error saving recent files: {str(e)}")
    
    def add_recent_file(self, file_path):
        """Add a file to recent files"""
        # Remove if already exists
        self.recent_files = [f for f in self.recent_files if f != file_path]
        
        # Add to beginning
        self.recent_files.insert(0, file_path)
        
        # Keep only the 5 most recent files
        self.recent_files = self.recent_files[:5]
        
        # Save and update UI
        self.save_recent_files()
        self.update_recent_files_ui()
    
    def update_recent_files_ui(self):
        """Update the recent files UI"""
        # Clear existing widgets
        for widget in self.files_frame.winfo_children():
            widget.destroy()
        
        # Add file links
        if not self.recent_files:
            no_files_label = ttk.Label(self.files_frame, 
                                    text="No recent files",
                                    background=ModernTheme.BG_COLOR)
            no_files_label.pack(anchor=tk.W, pady=2)
        else:
            for file_path in self.recent_files:
                file_name = os.path.basename(file_path)
                file_frame = ttk.Frame(self.files_frame)
                file_frame.pack(fill=tk.X, pady=2)
                
                # File icon
                file_icon = ttk.Label(file_frame, 
                                   text="📄",
                                   background=ModernTheme.BG_COLOR)
                file_icon.pack(side=tk.LEFT, padx=(0, 5))
                
                # File link
                file_link = tk.Label(file_frame, 
                                  text=file_name,
                                  fg=ModernTheme.ACCENT_COLOR,
                                  bg=ModernTheme.BG_COLOR,
                                  cursor="hand2")
                file_link.pack(side=tk.LEFT, fill=tk.X)
                
                # Bind click event
                file_link.bind("<Button-1>", lambda e, path=file_path: self.open_file(path))
                
                # Add tooltip with full path
                self.create_tooltip(file_link, file_path)
    
    def create_tooltip(self, widget, text):
        """Create a tooltip for a widget"""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25
            
            # Create a toplevel window
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            label = tk.Label(self.tooltip, text=text, 
                          background="#ffffe0", relief="solid", borderwidth=1,
                          font=("Segoe UI", 8))
            label.pack()
        
        def leave(event):
            if hasattr(self, "tooltip"):
                self.tooltip.destroy()
        
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
    
    def open_file(self, file_path):
        """Open a file with the default application"""
        try:
            if os.path.exists(file_path):
                if os.name == 'nt':  # Windows
                    os.startfile(file_path)
                else:  # macOS and Linux
                    subprocess.call(('xdg-open', file_path))
            else:
                messagebox.showerror("Error", f"File not found: {file_path}")
                # Remove from recent files
                self.recent_files = [f for f in self.recent_files if f != file_path]
                self.save_recent_files()
                self.update_recent_files_ui()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {str(e)}")
    
    def clear_credentials(self):
        """Clear saved login credentials"""
        app_data_dir = os.path.join(os.path.expanduser("~"), ".decopress")
        credentials_file = os.path.join(app_data_dir, "credentials.json")
        
        if os.path.exists(credentials_file):
            try:
                os.remove(credentials_file)
                messagebox.showinfo("Success", "Saved credentials have been cleared.\nYou will be prompted to log in next time.")
            except Exception as e:
                messagebox.showerror("Error", f"Could not clear credentials: {str(e)}")
        else:
            messagebox.showinfo("Information", "No saved credentials found.")
    
    def create_button(self, parent, text, command, is_primary=False):
        """Create a modern styled button"""
        btn = tk.Button(parent, text=text, command=command)
        
        # Apply button style
        for key, value in ModernTheme.BUTTON_STYLE.items():
            btn[key] = value
            
        # Adjust colors based on primary/secondary
        if not is_primary:
            btn['background'] = ModernTheme.SECONDARY_COLOR
            btn['activebackground'] = "#2d2d2f"
            
        return btn
        
    def run_daily_orders(self):
        """Run daily orders script with loading screen"""
        # Create a more visible loading screen
        try:
            # Show loading screen with a more direct approach
            loading_screen = LoadingScreen(self.root)
            
            # Process events to let UI update, but don't try to force focus
            for _ in range(5):  # Fewer update cycles
                try:
                    # Update window without forcing focus
                    loading_screen.top.update_idletasks()
                    
                    # Process UI events
                    self.root.update_idletasks()
                except Exception as e:
                    print(f"Error during initial loading screen update: {e}")
                    pass
            
            # Now run the task directly (not in a thread)
            try:
                result_file = None
                
                # Keep refreshing the loading screen while running the task
                def keep_alive():
                    try:
                        # Only update if the window still exists
                        if hasattr(loading_screen, 'top') and loading_screen.top.winfo_exists():
                            # Don't lift or try to take focus
                            try:
                                loading_screen.top.update()
                            except Exception as e:
                                print(f"Error updating loading screen: {e}")
                            # Schedule the next update
                            try:
                                loading_screen.top.after(100, keep_alive)
                            except Exception as e:
                                print(f"Error scheduling next keep_alive: {e}")
                    except Exception as e:
                        print(f"Error in keep_alive: {e}")
                        pass
                
                # Start the keep-alive function
                loading_screen.top.after(100, keep_alive)
                
                # Run daily orders
                result_file = daily_orders.run()
                
                # Add to recent files if successful
                if result_file and os.path.exists(result_file):
                    self.add_recent_file(result_file)
                    
                    # Close loading screen
                    loading_screen.close()
                    self.root.deiconify()
                    
                    # Show success message with filename
                    messagebox.showinfo("Success", 
                        f"Daily orders have been successfully exported!\n\nFile: {os.path.basename(result_file)}")
                    
                    # Open the file
                    self.open_file(result_file)
                else:
                    # Close loading screen
                    loading_screen.close()
                    self.root.deiconify()
                    
                    # Generic success message if no file was returned
                    messagebox.showinfo("Success", "Daily orders process completed.")
                
            except Exception as e:
                # Close loading screen
                try:
                    loading_screen.close()
                    self.root.deiconify()
                except:
                    pass
                
                # Show error
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
        except Exception as e:
            # If loading screen creation fails
            try:
                self.root.deiconify()  # Make sure main window is shown
            except:
                pass
            messagebox.showerror("Error", f"Could not create loading screen: {str(e)}")
    
    def run_packing_slip(self):
        """Run packing slip script with loading screen"""
        # Create a more visible loading screen
        try:
            # Show loading screen with a more direct approach
            loading_screen = LoadingScreen(self.root)
            
            # Process events to let UI update, but don't try to force focus
            for _ in range(5):  # Fewer update cycles
                try:
                    # Update window without forcing focus
                    loading_screen.top.update_idletasks()
                    
                    # Process UI events
                    self.root.update_idletasks()
                except Exception as e:
                    print(f"Error during initial loading screen update: {e}")
                    pass
            
            # Now run the task directly (not in a thread)
            try:
                # Keep refreshing the loading screen while running the task
                def keep_alive():
                    try:
                        # Only update if the window still exists
                        if hasattr(loading_screen, 'top') and loading_screen.top.winfo_exists():
                            # Don't lift or try to take focus
                            try:
                                loading_screen.top.update()
                            except Exception as e:
                                print(f"Error updating loading screen: {e}")
                            # Schedule the next update
                            try:
                                loading_screen.top.after(100, keep_alive)
                            except Exception as e:
                                print(f"Error scheduling next keep_alive: {e}")
                    except Exception as e:
                        print(f"Error in keep_alive: {e}")
                        pass
                
                # Start the keep-alive function
                loading_screen.top.after(100, keep_alive)
                
                result_files = None
                
                # Run packing slip
                result_files = packing_slip.run()
                
                # Add to recent files if successful
                if result_files:
                    excel_path, pdf_path = result_files
                    if excel_path and os.path.exists(excel_path):
                        self.add_recent_file(excel_path)
                    if pdf_path and os.path.exists(pdf_path):
                        self.add_recent_file(pdf_path)
                
                    # Close loading screen
                    loading_screen.close()
                    self.root.deiconify()
                    
                    # Show success message
                    messagebox.showinfo("Success", "Packing slip has been successfully created!")
                    
                    # Open the PDF if available, otherwise Excel
                    file_to_open = pdf_path if pdf_path and os.path.exists(pdf_path) else excel_path
                    if file_to_open:
                        self.open_file(file_to_open)
                else:
                    # Close loading screen
                    loading_screen.close()
                    self.root.deiconify()
                    
                    # Generic success message if no file was returned
                    messagebox.showinfo("Success", "Packing slip process completed.")
                
            except Exception as e:
                # Close loading screen
                try:
                    loading_screen.close()
                    self.root.deiconify()
                except:
                    pass
                
                # Show error
                messagebox.showerror("Error", f"An error occurred: {str(e)}")
        except Exception as e:
            # If loading screen creation fails
            try:
                self.root.deiconify()  # Make sure main window is shown
            except:
                pass
            messagebox.showerror("Error", f"Could not create loading screen: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DecoPressApp(root)
    root.mainloop() 