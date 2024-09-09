import tkinter as tk
from tkinter import messagebox
from threading import Thread
import subprocess
import sys
import os

# Function to get the correct path to the script files
def resource_path(relative_path):
    """ Get the absolute path to the resource, whether we're running as a script or as a PyInstaller bundle """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Paths to the scripts
# city_script_path = resource_path("Find Country of given city.py")
Taxonomy_SLN_script_path = resource_path("Find Taxonomy or SLN.py")
Find_NPI_from_website_script_path = resource_path("Find NPI from website.py")
Find_NPI_from_previous_year_script_path = resource_path("Find NPI from previous year file.py")

# Function to run the City script asynchronously
# def run_city_script():
#     update_status("Running City Finder...")
#     def run():
#         try:
#             subprocess.run(["python", city_script_path])
#             update_status("City Finder completed!")
#         except Exception as e:
#             messagebox.showerror("Error", f"Failed to run City script: {e}")
#             update_status("City Finder failed.")
#     Thread(target=run).start()

# Function to run the Taxonomy/SLN script asynchronously
def run_Taxonomy_SLN_script():
    update_status("Running Taxonomy/SLN Finder...")
    def run():
        try:
            subprocess.run(["python", Taxonomy_SLN_script_path])
            update_status("Taxonomy/SLN Finder completed!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run Taxonomy/SLN script: {e}")
            update_status("Taxonomy/SLN Finder failed.")
    Thread(target=run).start()

# Function to run the NPI from website script asynchronously
def run_NPI_from_website_script():
    update_status("Running NPI from Website Finder...")
    def run():
        try:
            subprocess.run(["python", Find_NPI_from_website_script_path])
            update_status("NPI from Website Finder completed!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run NPI from website script: {e}")
            update_status("NPI from Website Finder failed.")
    Thread(target=run).start()

# Function to run the NPI from previous year script asynchronously
def run_NPI_from_previous_year_script():
    update_status("Running NPI from Previous Year Finder...")
    def run():
        try:
            subprocess.run(["python", Find_NPI_from_previous_year_script_path])
            update_status("NPI from Previous Year Finder completed!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run NPI from previous year script: {e}")
            update_status("NPI from Previous Year Finder failed.")
    Thread(target=run).start()

# Function to update the status bar
def update_status(message):
    status_label.config(text=message)

# Function to change button style on hover
def on_enter(e):
    e.widget['background'] = 'black'
    e.widget['foreground'] = 'white'

def on_leave(e):
    e.widget['background'] = '#ffffff'
    e.widget['foreground'] = 'black'

# Create the main application window
app = tk.Tk()
app.title("City and Designation App")
app.geometry("500x400")  # Set the window size

# Set a modern flat background color
app.configure(bg="#f7f7f7")

# Add a title label with a modern look
title_label = tk.Label(app, text="NPI and Taxonomy Finder", font=("Helvetica", 18, "bold"), bg="#f7f7f7", fg="#333")
title_label.pack(pady=20)

# Create a frame to hold the content
frame = tk.Frame(app, bg="#f7f7f7")
frame.place(relx=0.5, rely=0.5, anchor="center", width=350, height=300)

# Create and place the buttons in the frame with a flat design and rounded corners
button_style = {"bg": "#ffffff", "fg": "black", "font": ("Helvetica", 12), "padx": 10, "pady": 10, "relief": "flat", "bd": 0}

# city_button = tk.Button(frame, text="City Finder", command=run_city_script, **button_style)
# city_button.pack(pady=10, ipadx=20)
# city_button.bind("<Enter>", on_enter)
# city_button.bind("<Leave>", on_leave)

Taxonomy_SLN_button = tk.Button(frame, text="Find Taxonomy/SLN from Website", command=run_Taxonomy_SLN_script, **button_style)
Taxonomy_SLN_button.pack(pady=10, ipadx=20)
Taxonomy_SLN_button.bind("<Enter>", on_enter)
Taxonomy_SLN_button.bind("<Leave>", on_leave)

Find_NPI_from_website_button = tk.Button(frame, text="Find NPI from Website", command=run_NPI_from_website_script, **button_style)
Find_NPI_from_website_button.pack(pady=10, ipadx=20)
Find_NPI_from_website_button.bind("<Enter>", on_enter)
Find_NPI_from_website_button.bind("<Leave>", on_leave)

NPI_from_previous_year_button = tk.Button(frame, text="Find NPI from Previous Year", command=run_NPI_from_previous_year_script, **button_style)
NPI_from_previous_year_button.pack(pady=10, ipadx=20)
NPI_from_previous_year_button.bind("<Enter>", on_enter)
NPI_from_previous_year_button.bind("<Leave>", on_leave)

# Add a status bar at the bottom
status_label = tk.Label(app, text="Ready", bd=1, relief="sunken", anchor="w", font=("Helvetica", 10), bg="#eeeeee", fg="#333")
status_label.pack(side="bottom", fill="x")

# Run the Tkinter event loop
app.mainloop()
