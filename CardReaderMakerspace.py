import sys
import customtkinter as ctk
from openpyxl import load_workbook
import tkinter as tk
from tkinter import simpledialog, Canvas, messagebox
import PIL
from PIL import Image, ImageTk 
from screeninfo import get_monitors
import pygetwindow as gw
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from pynput.keyboard import Key, Controller
import time
import requests
import os
import shutil
import threading


###    Version 1.2

# Path to excel sheet
file_path = "hardware_users.xlsx"
sheet_name = "Scans"
sheet2_name = "Users"
Location = "Watt"
keyboard = Controller()


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
                match_found = True
                print(f"User with hardware ID {hardware_id} already exists in 'Users' sheet.")
                break  # Stop searching after finding the match

            # If the hardware ID cell is empty (i.e., new entry row), fill in this row
            if cell_hardware_id is None or cell_hardware_id == "":
                row[0].value = username  # Column A for username
                row[1].value = int(hardware_id)  # Column B for hardware ID
                row[3].value = first_name  # Column D for first name
                row[4].value = last_name   # Column E for last name
                row[5].value = major       # Column F for major
                match_found = True
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
        return username, None, None  # Default values if element not found
    except Exception as e:
        print(f"Error during page load: {e}")
        driver.quit()
        return username, None, None  # Default values if there's an error

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
    if major_element:
        major = major_element.get_text().strip()
    else:
        major = None

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
    # Create new popup window as child of root
    popup = tk.Toplevel(root)
    popup.title("Welcome")
    
    # Get screen dimensions and set geometry
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    popup.geometry(f"{screen_width}x{screen_height}+0+0")
    
    # Configure popup window
    popup.attributes('-fullscreen', True)
    popup.attributes('-topmost', True)
    popup.focus_force()
    
    try:
        # Load and scale background image
        image = Image.open("backgroundLarge.png")
        image = image.resize((screen_width, screen_height), PIL.Image.Resampling.LANCZOS)
        bg_image = ImageTk.PhotoImage(image)
        popup.bg_image = bg_image  # Keep reference
        
        # Create background label
        background_label = tk.Label(popup, image=bg_image)
        background_label.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Set welcome message
        if first_name is None:
            first_name = username
        message = "Welcome to the Makerspace! Be sure to complete the safety trainings!" if userstatus == 1 else f"Welcome back, {first_name}!"
        
        # Create message label
        message_label = tk.Label(
            popup,
            text=message,
            font=("Helvetica", 40, "bold"),
            fg="white",
            bg="black"
        )
        message_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # Close after 2.5 seconds
        popup.after(2500, lambda: popup.destroy())
        
    except Exception as e:
        print(f"Error in welcome popup: {e}")
        popup.destroy()

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
    def is_valid_username(username):
        if username.isdigit():
            return False, "Username cannot be all numbers."
        if not username.isalnum():
            return False, "Username must be alphanumeric."
        return True, ""

def prompt_for_username():
    root = ctk.CTk()
    root.title("Username Entry")
    username = None
    
    # Configure window
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    root.geometry(f"{screen_width}x{screen_height}+0+0")
    root.attributes('-fullscreen', True)
    root.attributes('-topmost', True)
    
    def cleanup():
        try:
            root.quit()
            root.destroy()
        except Exception:
            pass
    
    def submit_username(event=None):
        nonlocal username
        entered_username = entry.get()
        if "@clemson.edu" in entered_username:
            entered_username = entered_username.split("@")[0]
            
        valid, message = is_valid_username(entered_username)
        if valid:
            username = entered_username
            root.after(100, cleanup)
        else:
            show_error_popup(message)
    
    # Create and configure widgets
    frame = ctk.CTkFrame(master=root)
    frame.pack(expand=True, fill="both")
    
    title_label = ctk.CTkLabel(
        master=frame,
        text="Welcome to the Makerspace! Enter Your Clemson Username:",
        font=("Arial", 55),
        text_color="#F56600"
    )
    title_label.pack(pady=20)
    
    entry = ctk.CTkEntry(
        master=frame,
        width=500,
        height=50,
        placeholder_text="Enter username",
        font=("Arial", 32)
    )
    entry.pack(pady=10)
    
    # Focus handling
    root.after(100, lambda: (
        entry.focus_force(),
        root.focus_force()
    ))
    
    # Bind events
    root.bind('<Return>', submit_username)
    root.bind('<Escape>', lambda e: cleanup())
    
    # Auto-close timer
    root.after(25000, cleanup)
    
    # Start main loop
    root.mainloop()
    
    return username

def scrape_user_thread(username, callback):
    def run():
        first_name, last_name, major = scrape_user(username)
        callback(first_name, last_name, major)
    thread = threading.Thread(target=run)
    thread.start()

def main():
    userstatus = None
    hardware_id = sys.argv[1]  # This gets the hardware ID from the global system variables as defined from the other script to pass along the variables.
    workbook, sheet, sheet2 = load_excel()
    username = find_hardware_id(sheet2, hardware_id)
    first_name = None
    major = None
    root = tk.Tk()
    root.withdraw()  # Hide the root window initially
    root.bind("<Escape>", close_on_escape)  # Bind the Escape key to close the program

    def on_scrape_complete(first_name, last_name, major):
        add_user_to_sheet(sheet_name, sheet2_name, hardware_id, username, first_name, last_name, major, workbook, userstatus)
        workbook.save(file_path)

    if username is not None:
        print(f"User found: {username}")
        userstatus = 0
        first_name, last_name, major = find_userdata(hardware_id, sheet2)
        add_user_to_sheet(sheet_name, sheet2_name, hardware_id, username, first_name, last_name, major, workbook, userstatus)
        show_welcome_popup(root, username, first_name, userstatus)
        root.mainloop()
    else:
        print("New user detected. Prompting for username.")
        userstatus = 1
        username = prompt_for_username()
        print(f'Username entered: {username}')
        if username is not None:
            # Show welcome popup immediately
            show_welcome_popup(root, username, None, userstatus)
            # Start scraping in background
            scrape_user_thread(username, on_scrape_complete)
            root.mainloop()
        else:
            print('Username Prompt timed out')
        username = None

if __name__ == "__main__":
    main()