import os 
import tkinter as tk
import PIL
from PIL import Image, ImageTk
from tkinter import Canvas
import random
import webbrowser
import subprocess
import time
import shutil
import threading
from datetime import datetime
import psutil
from openpyxl import load_workbook
import sys
from queue import Queue
import atexit  # For cleanup on program exit

# Import Excel ↔ Database bidirectional sync
try:
    from excel_db_sync import sync_on_startup, sync_on_shutdown
    BIDIRECTIONAL_SYNC_AVAILABLE = True
except ImportError:
    print("Warning: excel_db_sync module not found. Bidirectional sync disabled.")
    BIDIRECTIONAL_SYNC_AVAILABLE = False

# Import database sync (optional, based on config)
try:
    from database_sync import init_sync
    SYNC_AVAILABLE = True
except ImportError:
    SYNC_AVAILABLE = False

# Import SQLite database for username→hardware_id lookup
try:
    import database as db
    DB_AVAILABLE = True
except ImportError:
    print("Warning: database module not found. Username lookup disabled.")
    DB_AVAILABLE = False

def get_resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def cleanup_cardreader_flags():
    """Clean up any stale flag files from CardReaderMakerspace processes"""
    try:
        # Remove stale popup flag (used by CardReaderMakerspace.py)
        if os.path.exists(".popup_active"):
            flag_age = time.time() - os.path.getmtime(".popup_active")
            if flag_age > 10:  # Stale if older than 10 seconds
                os.remove(".popup_active")
                print(f"[CLEANUP] Removed stale CardReader popup flag ({flag_age:.1f}s old)")
        
        # Remove stale queued scan
        if os.path.exists(".queued_scan"):
            scan_age = time.time() - os.path.getmtime(".queued_scan")
            if scan_age > 30:  # Stale if older than 30 seconds
                os.remove(".queued_scan")
                print(f"[CLEANUP] Removed stale queued scan ({scan_age:.1f}s old)")
    except Exception as e:
        print(f"[CLEANUP] Error during flag cleanup: {e}")

Location = "Cooper"  # Change to "Watt" if needed
file_path = "hardware_users.xlsx"
sheet_name = "Scans"

# Queue system for handling scans while popup is open
scan_queue = Queue()
popup_active = False

# Check if there is enough storage space (e.g., at least 100MB free)
min_free_space_mb = 100
if psutil.disk_usage('/').free < min_free_space_mb * 1024 * 1024:
    raise Exception("Not enough storage space to run the script. Please free up some space and try again.")

def add_username_scan(username):
    """Add a username-only scan to the spreadsheet"""
    try:
        workbook = load_workbook(filename=file_path)
        scans_sheet = workbook[sheet_name]
        
        # Add the scan with empty hardware_id, username, and timestamp
        now = datetime.now()
        timestamp = now.strftime('%m/%d/%Y %H:%M:%S')
        scans_sheet.append(["", username, timestamp])  # Empty hardware_id for username-only scans
        
        workbook.save(file_path)
        print(f"Username scan added for {username}, workbook saved.")
        return True
    except Exception as e:
        print(f"Error adding username scan: {e}")
        return False

def show_generic_welcome_popup(username):
    """Show a generic welcome popup for username-only entries"""
    global popup_active
    
    popup_active = True
    popup = tk.Toplevel()
    popup.title("Welcome")
    
    # Get screen dimensions
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    
    # Configure popup window
    popup.attributes('-fullscreen', True)
    popup.attributes('-topmost', True)
    
    # Load and scale background image (location-specific)
    try:
        if Location == "Cooper":
            background_filename = "BackgroundAdobe.png"
        else:
            background_filename = "BackgroundWatt.png"
        
        image = Image.open(get_resource_path(background_filename))
        image = image.resize((screen_width, screen_height), PIL.Image.Resampling.LANCZOS)
        bg_image = ImageTk.PhotoImage(image)
        
        # Keep reference to prevent garbage collection
        popup.bg_image = bg_image
        
        # Create background label
        background_label = tk.Label(popup, image=bg_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
    except:
        # Fallback to solid color if image not found
        popup.configure(bg="black")
    
    # Create message label
    message = f"Welcome to the Makerspace, {username}!"
    message_label = tk.Label(popup, text=message, font=("Helvetica", 40, "bold"), fg="white", bg="black")
    message_label.place(relx=0.5, rely=0.5, anchor="center")
    
    # Close after 2.5 seconds and process queue
    def close_and_process_queue():
        global popup_active
        popup.destroy()
        popup_active = False
        process_scan_queue()
    
    popup.after(2500, close_and_process_queue)

def process_scan_queue():
    """Process any scans that were queued while popup was open"""
    global popup_active
    
    if not scan_queue.empty() and not popup_active:
        user_input = scan_queue.get()
        print(f"Processing queued scan: {user_input}")
        process_scan_input(user_input)

def process_scan_input(user_input):
    """Process a scan input (either from direct entry or queue)"""
    global popup_active
    
    # Check if the input is exactly 6 digits and numerical
    if user_input.isdigit() and len(user_input) == 6:
        hardware_id = user_input
        print(f"Hardware ID entered: {hardware_id}")
        # Use the correct path for the CardReaderMakerspace script
        card_reader_path = get_resource_path("CardReaderMakerspace.py")
        subprocess.Popen([sys.executable, card_reader_path, hardware_id])
    else:
        # Treat as username
        username = user_input.strip()
        if username:  # Only process if not empty
            print(f"Username entered: {username}")
            
            # Try to look up the username in the database to get their hardware_id
            # If found, route through CardReaderMakerspace for the full popup (with training data)
            if DB_AVAILABLE:
                try:
                    user = db.get_user_by_username(username)
                    if user and user.get('hardware_id'):
                        hardware_id = str(user['hardware_id'])
                        print(f"Username '{username}' found in database with HWID {hardware_id}. Routing through CardReader.")
                        # Record the scan in database and Excel
                        now = datetime.now()
                        timestamp = now.strftime('%m/%d/%Y %H:%M:%S')
                        try:
                            db.add_scan(hardware_id, username, timestamp, Location)
                        except Exception as e:
                            print(f"Database scan record error: {e}")
                        add_username_scan(username)
                        # Launch CardReaderMakerspace with the hardware_id for full popup
                        card_reader_path = get_resource_path("CardReaderMakerspace.py")
                        subprocess.Popen([sys.executable, card_reader_path, hardware_id])
                        return
                except Exception as e:
                    print(f"Database lookup error for username '{username}': {e}")
            
            # Fallback: user not found in database - add as username-only scan
            if add_username_scan(username):
                # Show welcome popup
                show_generic_welcome_popup(username)

# Function to handle input from the text box
def handle_entry(event=None):
    global popup_active
    
    user_input = entry.get()
    entry.delete(0, tk.END)  # Clear the entry box immediately
    
    # If popup is active, queue the scan
    if popup_active:
        print(f"Popup active - queueing scan: {user_input}")
        scan_queue.put(user_input)
    else:
        # Process immediately
        process_scan_input(user_input)

# Employee Clock-In button, change this link when we move away from kronos.
def open_clock_in():
    webbrowser.open("https://clemson.kronos.net")

# Create the main window
root = tk.Tk()
root.title("Sign In")
root.attributes('-fullscreen', True)  # Make it fullscreen

# Create a canvas for the gradient background
canvas = Canvas(root, width=root.winfo_screenwidth(), height=root.winfo_screenheight())
canvas.pack(fill="both", expand=True)

# Load the image and display it below the entry box (location-specific)
if Location == "Cooper":
    image_path = get_resource_path("BackgroundAdobe.png")
else:
    image_path = get_resource_path("BackgroundTablet.png")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
image = Image.open(image_path)
image = image.resize((screen_width, screen_height), PIL.Image.Resampling.LANCZOS)
image = ImageTk.PhotoImage(image)

# Create a label to hold the image and place it below the entry box
image_label = tk.Label(canvas, image=image, bg='#F56600')
image_label.place(x=0, y=0, relwidth=1, relheight=1)

# Text instead of adding it as a part of the background image, not using due to lack of customiation in tkinter.
#text_label = tk.Label(canvas, text="Scan the reader to sign in", font=("Vendeta", 60), bg='#F56600', fg='#522D80')
#text_label.place(relx=0.5, rely=0.3, anchor='center')  # Place text label in the middle

# Choose accent color based on location
if Location in ("Cooper", "Adobe"):
    accent_color = '#CC0000'  # Red for Cooper/Adobe
else:
    accent_color = '#F56600'  # Clemson orange for other locations

# Create a text entry box with rounded corners and border
entry_frame = tk.Frame(canvas, bg=accent_color, bd=3, highlightbackground=accent_color, highlightcolor=accent_color, highlightthickness=3)
entry_frame.place(relx=0.5, rely=0.7, anchor='center')

# Apply modern styling to the Entry widget
entry = tk.Entry(entry_frame, font=("Vendeta", 30), justify='center', bd=0, relief=tk.FLAT)
entry.config(bg='white', fg='#333333', insertbackground=accent_color, highlightthickness=2, highlightbackground=accent_color, highlightcolor=accent_color)
entry.pack(ipadx=10, ipady=5, padx=10, pady=5)  # Padding for a modern look
entry.focus_set()  # Focus the text box automatically

# ClockIn button (Optional and lowkey useless with kronos rn, might remove, people just use their phones)
clock_in_button = tk.Button(canvas, text="Employee Clock-In", font=("Helvetica", 16), bg='#522D80', fg='white', command=open_clock_in)
clock_in_button.place(x=10, y=10)

# Confetti animation
confetti_items = []

def create_confetti():
    """Create small rectangles to represent confetti falling from the top of the screen."""
    for _ in range(100):  # Create 100 confetti pieces
        x_position = random.randint(0, root.winfo_screenwidth())
        y_position = random.randint(-root.winfo_screenheight(), 0)  # Start off-screen
        size = random.randint(5, 15)
        color = random.choice(['#F56600', '#522D80', '#FFDD00', '#00FFDD', '#FF66CC'])
        confetti = canvas.create_rectangle(x_position, y_position, x_position + size, y_position + size, fill=color, outline=color)
        confetti_items.append((confetti, random.randint(2, 10)))  # Assign a random fall speed

def animate_confetti():
    """Animate the confetti falling down the screen."""
    for confetti, speed in confetti_items:
        canvas.move(confetti, 0, speed)  # Move downwards by the speed value
            
    root.after(1, animate_confetti)  # Continue the animation

def start_confetti():
    """Start the confetti creation and animation."""
    create_confetti()
    animate_confetti()
    root.after(1000, stop_confetti)  # Let confetti play for however many seconds

def stop_confetti():
    """Clear all confetti after the duration is over."""
    for confetti, _ in confetti_items:
        canvas.delete(confetti)
    confetti_items.clear()

# Backup functionality
def backup_file():
    while True:
        try:
            now = datetime.now()
            date_str = now.strftime("%m_%d_%Y___%H_%M_%S")
            backup_folder = "backups"
            if not os.path.exists(backup_folder):
                os.makedirs(backup_folder)
            backup_filename = f"hardware_users_{Location}_{date_str}.xlsx"
            backup_path = os.path.join(backup_folder, backup_filename)
            shutil.copy("hardware_users.xlsx", backup_path)
            print(f"Backup created: {backup_path}")
        except Exception as e:
            print(f"Failed to create backup: {e}")
        time.sleep(86400)  # Wait for 24 hours (86400 seconds)

# Perform bidirectional Excel ↔ Database sync on startup
if BIDIRECTIONAL_SYNC_AVAILABLE:
    try:
        sync_on_startup()
        # Register shutdown sync to run when program exits
        atexit.register(sync_on_shutdown)
    except Exception as e:
        print(f"Note: Bidirectional sync failed: {e}")

# Initialize database sync system (if enabled)
if SYNC_AVAILABLE:
    try:
        init_sync()
    except Exception as e:
        print(f"Note: Database sync initialization failed: {e}")

# Clean up any stale CardReader flag files on startup
print("Checking for stale CardReader flag files...")
cleanup_cardreader_flags()

# Start the backup thread
backup_thread = threading.Thread(target=backup_file, daemon=True)
backup_thread.start()

# Bind the Enter key to trigger storing the input and confetti animation
entry.bind('<Return>', handle_entry)

# Exit the program when pressing 'Esc'
root.bind('<Escape>', lambda e: root.destroy())

root.mainloop()
