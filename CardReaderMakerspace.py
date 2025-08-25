import sys
import customtkinter as ctk
from openpyxl import load_workbook
import tkinter as tk
import PIL
from PIL import Image, ImageTk 
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import threading


###    Version 1.2

# Path to excel sheet
file_path = "hardware_users.xlsx"
sheet_name = "Scans"
sheet2_name = "Users"
Location = "Watt"


def load_excel():
    # Load the workbook and sheet
    workbook = load_workbook(filename=file_path)
    sheet = workbook[sheet_name]
    sheet2 = workbook[sheet2_name]
    return workbook, sheet, sheet2

def find_hardware_id(sheet2, hardware_id):
    # Loop through the rows starting from row 2 to skip headers (if any)
    for row in sheet2.iter_rows(min_row=2, max_row=sheet2.max_row, values_only=True):
        if str(row[1]) == str(hardware_id):  # Compare hardware_id in column B (index 1)
            return row[0]  # Return the username from column A (index 0)
    return None  # Return None if not found

def find_userdata(hardware_id, sheet2):
    # Loop through the rows starting from row 2 to skip headers
    for row in sheet2.iter_rows(min_row=2, max_row=sheet2.max_row, values_only=True):
        if str(row[1]) == str(hardware_id):  # Look for hardware_id in col. B
            first_name = row[3]  # Column D 
            last_name = row[4]   # Column E 
            major = row[5]       # Column F
            return first_name, last_name, major
    return None, None, None  # Set to None if they are not found

def add_user_to_sheet(sheet_name, sheet2_name, hardware_id, username, first_name, last_name, major, workbook, userstatus):
    wb = workbook
    scans_sheet = wb[sheet_name]
    users_sheet = wb[sheet2_name]
    
    if userstatus == 1:
        # Search for matching hardware ID in the "Users" sheet or for an empty hardware ID cell
        for row in users_sheet.iter_rows(min_row=2, values_only=False):  # Skip header row
            cell_hardware_id = row[1].value  # Column B in "Users" for hardware_id

            # Cast both the hardware_id from input and the one from the sheet to str for comparison
            if str(cell_hardware_id) == str(hardware_id):
                print(f"User with hardware ID {hardware_id} already exists in 'Users' sheet.")
                break  # Stop searching after finding the match

            # If the hardware ID cell is empty (i.e., new entry row), fill in this row
            if cell_hardware_id is None or cell_hardware_id == "":
                row[0].value = username  # Column A for username
                row[1].value = int(hardware_id)  # Column B for hardware ID
                row[3].value = first_name  # Column D for first name
                row[4].value = last_name   # Column E for last name
                row[5].value = major       # Column F for major
                print(f"New user {first_name} {last_name} added to 'Users' sheet.")
                break  # Stop searching after appending the new data

    # Add the scan to the 'Scans' sheet (this happens regardless of userstatus)
    now = datetime.now()
    timestamp = now.strftime('%m/%d/%Y %H:%M:%S')  # Format the time to display as "YYYY-MM-DD HH:MM"
    scans_sheet.append([int(hardware_id), username, timestamp])
    
    # Save the workbook after making changes
    wb.save(file_path)
    print(f"Scan Added, workbook saved.")

def scrape_user(username):
    # Set up Selenium with headless Chrome
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run Chrome in headless mode
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    # Set up the driver
    driver = webdriver.Chrome(options=chrome_options)

    # Create the URL using username
    url = f"https://my.clemson.edu/#/directory/person/{username}"

    # Load the page
    driver.get(url)

    # Wait for the full name element to appear
    try:
        WebDriverWait(driver, 1).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.personView .primaryInfo h2'))
        )
    except TimeoutException:
        print(f"Timeout while waiting for the element on the page for {username}.")
        driver.quit()
        return None, None, None
    except Exception as e:
        print(f"Error during page load: {e}")
        driver.quit()
        return None, None, None

    # Get the page
    page_source = driver.page_source

    # Parse the page with BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')

    # Find the Name
    full_name_element = soup.select_one('.personView .primaryInfo h2')
    if full_name_element:
        full_name = full_name_element.get_text().strip()
        name_parts = full_name.split()
        first_name = name_parts[0]  # The first part of the name
        last_name = name_parts[-1]  # The last part of the name
    else:
        first_name, last_name = None, None

    # Find major
    major_element = soup.select_one('.personView .primaryInfo .data p')
    major = major_element.get_text().strip() if major_element else None

    # Print the scraped information
    print(f"First Name: {first_name}")
    print(f"Last Name: {last_name}")
    print(f"Major: {major}")

    # Close the driver
    driver.quit()
    return first_name, last_name, major

def make_fullscreen_on_top(root):
    root.attributes('-fullscreen', True)
    root.attributes('-topmost', True)

def show_welcome_popup(root, username, first_name, userstatus):
    popup = tk.Toplevel()
    popup.title("Welcome")
    
    # Get screen dimensions
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    
    # Configure popup window
    popup.attributes('-fullscreen', True)
    popup.attributes('-topmost', True)
    
    # Load and scale background image
    image = Image.open("backgroundLarge.png")
    image = image.resize((screen_width, screen_height), PIL.Image.Resampling.LANCZOS)
    bg_image = ImageTk.PhotoImage(image)
    
    # Keep reference to prevent garbage collection
    popup.bg_image = bg_image
    
    # Create background label
    background_label = tk.Label(popup, image=bg_image)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)
    
    # Set welcome message
    if first_name is None:
        first_name = username
    message = "Welcome to the Makerspace!" if userstatus == 1 else f"Welcome back, {first_name}!"
    
    # Create message label
    message_label = tk.Label(popup, text=message,font=("Helvetica", 40, "bold"), fg="white", bg="black")
    message_label.place(relx=0.5, rely=0.5, anchor="center")
    
    # Close after 2.5 seconds
    popup.after(2500, popup.destroy)

def close_on_escape(event):
    print("Escape key pressed. Exiting program...")
    sys.exit()  # Exit the program

def show_error_popup(message):
    # Create a new window for the error popup
    error_root = tk.Tk()
    error_root.title("Error")
    error_root.attributes('-fullscreen', True)  # Make it fullscreen

    # Set the background color
    error_root.configure(bg="black")

    # Create a label to display the error message
    error_label = tk.Label(error_root, text=message, font=("Helvetica", 40, "bold"), fg="white", bg="black")
    error_label.place(relx=0.5, rely=0.5, anchor="center")  # Center the message

    # Close the popup after 3 seconds
    error_root.after(3000, error_root.destroy)

    # Start the application loop
    error_root.mainloop()

def submit_username(event, entry, root):
    global username
    username = entry.get().strip()
    root.quit()

def prompt_for_username():
    global username
    username = None
    
    # Initialize the main window
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Username Entry")
    
    # Get screen dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Set window geometry and make it cover the entire screen
    root.geometry(f"{screen_width}x{screen_height}+0+0")
    root.resizable(False, False)
    root.attributes('-topmost', True)
    
    # Force the window to update and then configure it properly
    root.update_idletasks()
    root.lift()
    root.focus_force()
    
    # Create a centered frame for content (don't use pack with expand for main container)
    center_frame = ctk.CTkFrame(master=root, width=800, height=400)
    center_frame.place(relx=0.5, rely=0.5, anchor="center")

    # Title label
    title_label = ctk.CTkLabel(master=center_frame, text="Welcome to the Makerspace!\nEnter Your Clemson Username:", 
                              font=("Arial", 48), text_color="#F56600")
    title_label.pack(pady=30, padx=50)

    # Instruction label
    label = ctk.CTkLabel(master=center_frame, text="The part before the @clemson.edu", 
                        font=("Arial", 32))
    label.pack(pady=10, padx=50)

    # Create an entry box
    entry = ctk.CTkEntry(master=center_frame, width=600, height=60, 
                        placeholder_text="Enter username", font=("Arial", 28))
    entry.pack(pady=20, padx=50)
    
    # Function to check if input is a 6-digit hardware ID
    def check_hardware_id_input(event=None):
        global username, pending_hardware_id
        current_text = entry.get().strip()
        
        # If exactly 6 digits are entered, treat as hardware ID scan
        if current_text.isdigit() and len(current_text) == 6:
            pending_hardware_id = current_text
            username = None
            root.quit()
            return
        
        # Otherwise treat as username submission
        if event and event.keysym == 'Return':
            submit_username(event, entry, root)
    
    # Monitor text changes for hardware ID detection
    entry.bind('<KeyRelease>', check_hardware_id_input)
    
    # Bind escape key to close
    def on_escape(event):
        global username
        username = None
        root.quit()
    
    root.bind('<Escape>', on_escape)
    
    # Bind the Enter key to submit the form
    root.bind('<Return>', check_hardware_id_input)
    
    # Force focus to the entry box with multiple attempts
    def set_focus():
        root.focus_force()
        root.lift()
        entry.focus_set()
        entry.focus_force()
        root.after(50, lambda: entry.focus_force())
    
    root.after(100, set_focus)
    root.after(300, set_focus)  # Try again after longer delay
    
    # Auto-close after timeout
    def timeout_close():
        global username
        username = None
        root.quit()
    
    root.after(25000, timeout_close)
    
    root.mainloop()
    root.destroy()
    return username

def scrape_user_thread(username, callback):
    def run():
        first_name, last_name, major = scrape_user(username)
        callback(first_name, last_name, major)
    thread = threading.Thread(target=run)
    thread.start()

def main():
    global pending_hardware_id
    pending_hardware_id = None
    
    hardware_id = sys.argv[1]
    workbook, sheet, sheet2 = load_excel()
    username = find_hardware_id(sheet2, hardware_id)
    
    root = tk.Tk()
    root.withdraw()
    root.bind("<Escape>", close_on_escape)

    def on_scrape_complete(first_name, last_name, major):
        add_user_to_sheet(sheet_name, sheet2_name, hardware_id, username, first_name, last_name, major, workbook, 1)

    if username is not None:
        print(f"User found: {username}")
        first_name, last_name, major = find_userdata(hardware_id, sheet2)
        add_user_to_sheet(sheet_name, sheet2_name, hardware_id, username, first_name, last_name, major, workbook, 0)
        show_welcome_popup(root, username, first_name, 0)
        root.mainloop()
    else:
        print("New user detected. Prompting for username.")
        username = prompt_for_username()
        print(f'Username entered: {username}')
        
        # Check if a hardware ID was scanned during username entry
        if pending_hardware_id is not None:
            print(f"Hardware ID {pending_hardware_id} scanned during username entry. Processing new scan...")
            # Process the new hardware ID scan
            import subprocess
            subprocess.Popen(["python", "CardReaderMakerspace.py", pending_hardware_id])
            return
        
        if username:
            show_welcome_popup(root, username, None, 1)
            scrape_user_thread(username, on_scrape_complete)
            root.mainloop()
        else:
            print('Username Prompt timed out')

if __name__ == "__main__":
    main()