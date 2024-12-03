# This is an application that will read into an excel file from the input_files folder and take 
# and paste specific write specific data into the excel file in the output_files folder, all this
# will be done via a user interface built using Tkinter.

import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox

# Function to read the Excel file
def extract_data(input_file, output_file, columns_to_extract):
    pass

# Create the GUI using Tkinter
root = tk.Tk()
root.title("Customer Cargo Position Extractor")
# Set the window size
root.geometry("400x200")

# Add UI elements (buttons, labels, etc.)
input_file_label = tk.Label(root, text="Import File:")
input_file_label.pack(pady=5)

input_file_entry = tk.Entry(root, width=40)
input_file_entry.pack(pady=5)

# Button to browse for input file
def browse_input_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    input_file_entry.delete(0, tk.END) # Clear current entry
    input_file_entry.insert(0, filename) # Insert selected file path

browse_button = tk.Button(root, text="Browse", command=browse_input_file)
browse_button.pack(pady=5)

# Specify output file and columns to extract (We will need to work on this)
output_file = "path"
columns_to_extract = ["column1","column2"]

# Button to extract data
extract_button = tk.Button(root, text="Extract Data", command=lambda: extract_button_click())
extract_button.pack(pady=10)


# Add event handlers for button clicks, etc.
def extract_button_click():
    input_file = input_file_entry.get()
    if not input_file:
        messagebox.showwarning("Input Error", "Please select an input file.")
        return

    extract_data(input_file, output_file, columns_to_extract)

# Start the GUI event loop
root.mainloop()
