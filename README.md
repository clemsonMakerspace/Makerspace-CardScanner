# Makerspace Card Reader Script

This repository contains the **Makerspace Card Reader Script** for managing user check-ins and retrieving user data from the student directory. The script uses an Excel document to log user details and scan-ins, making it easy to keep track of makerspace usage.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Excel Setup](#excel-setup)
- [Running the Script](#running-the-script)
- [Libraries](#libraries)

## Overview
- The Card Reader Script was made to simplify sign-in to the makerspace and replace the web-interface given on the sign-in tablet ([visit.cumaker.space](https://visit.cumaker.space/))
- The goal of the script is to associate scan data (hardware ID) to a clemson user account and record time scanned.

## Features
- Automatic scan-in and logging of makerspace users with only one username entry per scan device.
- Integration with student directory for retrieving user data.
- Data logging in an Excel file.

## Requirements

- **Python 3.x** installed.
     It is reccomended to check the option "Add to PATH" and "Add Python to Enviroment Variables" in the installer options.
- Navigate to the cloned folder including "Makerspace-CardScanner" and run the following command in command prompt.
```bash
     pip install -r requirements.txt
```
- **Chrome Web Driver**installed: should be installed by default if you have crome downloaded or are on windows generally. 

## Excel Setup
1. **Excel Document**:  
   The script requires an Excel document named `hardware_users.xlsx` to store user details and scan-ins. This document should contain:
   - Two sheets:
     - **Scans**: To log the scan-ins (time and hardware_id).
     - **Users**: To store user details such as name, ID, major, and training data (to be added).
   - The sheet names and file name can be adjusted by modifying the first few variables in the script `CardReaderMakerspace.py`.
   - Before writing of the script will work, please enter an example entry into row 2 and a heading into row 1 of "Users".

2. **Script Setup**:
   - Clone this repository.
   - Ensure you have "background.png" and "BackgroundTablet.png" in main directory. #will be included in the github at a later date.
   - Install the required Python libraries using:
     ```bash
     pip install -r requirements.txt
     ```
   - Ensure the Excel file `hardware_users.xlsx` is not open when running the script, as the data will not be able to write if it's locked by another process.

## Running the Script
To run the card reader script:
1. Start the Python script `MakerspaceSignInTablet.py` by executing the following in the terminal:
   ```bash
   python MakerspaceSignInTablet.py
## Libraries 
An explanation of the required Libraries:

- **Tkinter**: Built-in Python library for GUI creation.
- **Random**: Part of Python’s standard library for generating random numbers. This was going to be used for random chance of fun popups or music to keep the script engaging for interns.
- **Webbrowser**: Standard Python library to open web browsers.
- **Subprocess**: Standard library for running external scripts/commands.
- **Sys**: For interacting with system-specific functions.
- **Pillow (PIL)**: Used for image manipulation.
- **CustomTkinter**: Used for modern basic graphics.
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
