# Makerspace Card Reader Script

The **Makerspace Card Reader Script** manages user check-ins and retrives user data from the clemson student directory to form its own database. The script uses an Excel document to log user details and scan-ins, making it easy to keep track of makerspace usage.

![MakerspaceSignInTablet_wVmxfdiXdg](https://github.com/user-attachments/assets/15e0cceb-7bd7-4f53-afce-c2dc29722155)
-----------------------------------------------------------------------------------
## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Excel Setup](#excel-setup)
- [Running the Script](#running-the-script)
- [Libraries](#libraries)
- [Troubleshooting](#troubleshooting)

## Overview
- The Card Reader Script was made to simplify sign-in to the makerspace and replace the web-interface given on the sign-in tablet ([visit.cumaker.space](https://visit.cumaker.space/))
- The goal of the script is to associate scan data (hardware ID) to a clemson user account and record a time stamp for the scan.

## Features
- Automatic scan-in and logging of makerspace users with only one username entry per scan device.
- Integration with student directory for retrieving user data.
- Data logging and editing in an Excel file.
- Automated Backups: Upon opening the script and every 24 hours, the script will save an automated backup to the backups folder. There is currently not a limit to the amount of backups so ensure to add this when implementing on any device with small avalible memory.

## Requirements

- **Python 3.x** installed.
     It is reccomended to check the option "Add to PATH" and "Add Python to Enviroment Variables" in the installer options.
- Navigate to the cloned folder including "Makerspace-CardScanner" and run the following command in command prompt.
```bash
     pip install -r requirements.txt
```
- **Chrome Web Driver** : should be installed already if you have chrome or a chromium based browser installed.
- **Excel or Alternative** : Should be able to edit an .xlsx document.

## Excel Setup
1. **Excel Document**:  
   The script requires an Excel document named `hardware_users.xlsx` to store user details and scan-ins. This document should contain:
   - Two sheets:
     - **Scans**: To log the scan-ins (time and hardware_id).
     - **Users**: To store user details such as name, ID, major, and training data (to be added).
   - The sheet names and file name can be adjusted by modifying the first few variables in the script `CardReaderMakerspace.py`.
   - Before writing of the script will work, please enter an example entry into row 2 and a heading into row 1 of "Users". (The data does not matter, just ensure the sheet is not blank.)

2. **Script Setup**:
   - Clone this repository.
   - Ensure you have "background.png" and "BackgroundTablet.png" in main directory.
   - Install the required Python libraries using:
     ```bash
     pip install -r requirements.txt
     ```
     ***IMPORTANT***: This is currently not functional as the requirements txt is not up to date. Currently you need to keep installing the missing libraries manually. I am sorry in advanced.
   - Ensure the Excel file `hardware_users.xlsx` is not open when running the script, as the data will not be able to write if it's locked by another process.
   - If script still does not run, ensure to pip install remaining missing libraries until program will run, this may take a while.

## Running the Script
To run the card reader script:
1. Start the Python script `MakerspaceSignInTablet.py` by executing the following in the terminal:
   ```bash
   python MakerspaceSignInTablet.py
## Libraries 
An explanation of the required Libraries:

- **Tkinter**: Built-in Python library for GUI.
- **Random**: Part of Python’s standard library for generating random numbers. This was going to be used for random chance of fun popups or music to keep the script engagement but has not been implemented yet.
- **Webbrowser**: Standard Python library to open web browsers and used to open student directory in background.
- **Subprocess**: Standard library for running external scripts/commands. The second .py file is opened this way to show over the scan screen.
- **Sys**: For interacting with system-specific functions.
- **Pillow (PIL)**: Used for image manipulation.
- **CustomTkinter**: Used for modern basic graphics. Currently only used for second script.
- **OpenPyXL**:For working with Excel documents.
- **Screeninfo**: To get monitor display information to adjust the script size based on the display to make it work fullscreen on any computer.
- **PyGetWindow**: Used for window management tasks. I don't remember why this is in there.
- **Datetime**: Part of Python’s standard library for working with date and time.
- **Selenium**: Used for automating web browser interactions (to access the student directory)
  - **Selenium WebDriver**: For controlling the web browser.
  - **Selenium Chrome Options**: For configuring Chrome browser options.
  - **Selenium Support UI**: For waiting for specific conditions to be met.
  - **Selenium By**: For locating elements on a webpage.
  - **Selenium Exceptions**: To handle browser-specific errors.
- **BeautifulSoup**: for retrieving user data from the student directory.
- **Internet access** for retrieving user data from the student directory.

## Troubleshooting:
- Within the first year of implementation the most common error seems to be the corruption of the excel file during the write operation. I am unsure what causes this but its likely to do with multiple scans writing during the save operation of the excel sheet. Do not use the "fixed or recovered" excel sheets as they do not work, you need to replace the hardware_users.xlsx entirely with a functional excel file.
 **Solution**: The script makes automated backups upon the launch and every 24 hours, please locate the most recent functional excel sheet and replace the "hardware_users.xlsx" with a functional file by renaming and swapping the files.


Donations: Paw points for a Caniac Combo are accepted. (Please please please please)
