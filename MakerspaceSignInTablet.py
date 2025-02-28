import tkinter as tk #graphics
import PIL
from PIL import Image, ImageTk #graphics
from tkinter import Canvas# also graphics
import random #Plans to add randomized backgrounds or random chance events on scan.
import webbrowser #to open browser
import subprocess #to open other script

# Function to handle input from the text box
def handle_entry(event=None):
    global hardware_id, username
    user_input = entry.get()
    
    # Check if the input is exactly 6 digits and numerical
    if user_input.isdigit() and len(user_input) == 6:
        hardware_id = user_input
        print(f"Hardware ID entered: {hardware_id}")
        username = None
        #start_confetti()  # YIPPPIEEEE CONFETTTIIIIII (commented out bc not being used rn)

        if 'hardware_id' in globals():
            subprocess.Popen(["python", "CardReaderMakerspace.py", hardware_id])
        elif 'username' in globals():
            subprocess.Popen(["python", "CardReaderMakerspace.py", username])
        else:
            print("No hardware_id or username is defined.")
    else:
        username = user_input
        print(f"Username entered: {username}")
    
    entry.delete(0, tk.END)  # Clear the entry box

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

# Load the image and display it below the entry box
image_path = "BackgroundTablet.png"  # file path
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

# Create a text entry box with rounded corners and border
entry_frame = tk.Frame(canvas, bg='#F56600', bd=0)
entry_frame.place(relx=0.5, rely=0.5, anchor='center')

# Apply modern styling to the Entry widget
entry = tk.Entry(entry_frame, font=("Vendeta", 30), justify='center', bd=0, relief=tk.FLAT)
entry.config(bg='white', fg='#333333', insertbackground='#522D80', highlightthickness=1, highlightbackground='#522D80', highlightcolor='#522D80')
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

# Bind the Enter key to trigger storing the input and confetti animation
entry.bind('<Return>', handle_entry)

# Exit the program when pressing 'Esc'
root.bind('<Escape>', lambda e: root.destroy())

root.mainloop()
