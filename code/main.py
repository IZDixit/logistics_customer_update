# This is an application that will read into an excel file from the input_files folder and take 
# and paste specific write specific data into the excel file in the output_files folder, all this
# will be done via a user interface built using Tkinter.

import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox,Frame
from datetime import datetime

# Function to read the Excel file
def extract_data(input_file, output_file, columns_to_extract):
    pass

# Specify output file and columns to extract (We will need to work on this)
def columns_extracted():
    customer_lists = ["default","customer1","customer2"]
    # We will need to add a condition based on who the customer is that will decide which path to take.
    output_file = "path"
    columns_to_extract = ["column1","column2"]
class ScrollableFrame(tk.Frame):
    """ A scrollable frame that can contain other widgets """
    def __init__(self, parent):
        super().__init__(parent)

        # Create a canvas and scrollbar
        self.canvas = tk.Canvas(self)
        self.scrollbar = tk.Scrollbar(self, orient="vertical",command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas)

        # Add the frame to the canvas with a center anchor
        self.canvas_window = self.canvas.create_window((0,0), window=self.scrollable_frame, anchor="n")

        # Configure the canvas to resize the window and keep it centered
        self.scrollable_frame.bind("<Configure>", lambda e:self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._center_window)        

        # Pack the canvas and scrollbar
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

    def _center_window(self, event):
        """ Center the scrollable frame within the canvas """
        canvas_width = event.width
        frame_width = self.scrollable_frame.winfo_reqwidth()
        x_offset = (canvas_width - frame_width) // 2
        self.canvas.coords(self.canvas_window, x_offset, 0)
class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Customer Cargo Position Extractor")
        # Set the window size
        self.geometry("400x300")

        # Create a scrollable Frame
        scrollable_frame = ScrollableFrame(self)
        scrollable_frame.pack(fill=tk.BOTH, expand=True)

        # Input File Selection
        input_file_label = tk.Label(scrollable_frame.scrollable_frame, text="Import File:")
        input_file_label.pack(pady=5)

        # Browse input file
        browse_button = tk.Button(scrollable_frame.scrollable_frame, text="Browse", command=lambda: browse_input_file(self))
        browse_button.pack(pady=5)

        self.input_file_entry = tk.Entry(scrollable_frame.scrollable_frame, width=40)
        self.input_file_entry.pack(pady=5)


        # Select the customer report is being extracted for
        customer_label = tk.Label(scrollable_frame.scrollable_frame, text="Select Customer:")
        customer_label.pack(pady=5)

        # Customer Selection
        self.customer_var = tk.StringVar(value="default")
        customer_menu = tk.OptionMenu(scrollable_frame.scrollable_frame, self.customer_var, "default","customer1","customer2")
        customer_menu.pack(pady=5)


        # Select date that will filter which rows of the above selected columns will be extracted
        date_label = tk.Label(scrollable_frame.scrollable_frame, text="Select Date (YYYY-MM-DD):")
        date_label.pack(pady=5)

        # Create a frame to hold the drop-down menus
        date_frame = tk.Frame(scrollable_frame.scrollable_frame)
        date_frame.pack(pady=10,anchor="center")


        # Create drop-downs for day, month, year
        days = list(range(1, 32))
        months = list(range(1,13))
        years = list(range(datetime.now().year - 3, datetime.now().year + 2)) # Last 3 years next 2 years

        self.day_var = tk.StringVar(value=str(datetime.now().day))
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.year_var = tk.StringVar(value=str(datetime.now().year))

        year_menu = tk.OptionMenu(date_frame, self.year_var, *years)
        year_menu.pack(side=tk.LEFT, padx=5)

        month_menu = tk.OptionMenu(date_frame, self.month_var, *months)
        month_menu.pack(side=tk.LEFT, padx=5)

        # Create a custom dropdown for days
        day_menu_button = tk.Menubutton(date_frame, textvariable=self.day_var, relief=tk.RAISED)
        day_menu_button.pack(side=tk.LEFT ,padx=5)

        # Create the menu
        day_menu = tk.Menu(day_menu_button, tearoff=0)
        day_menu_button.config(menu=day_menu)

        # Add two submenus for 1-15 and 16-31
        day_menu_column1 = tk.Menu(day_menu, tearoff=0)
        day_menu_column2 = tk.Menu(day_menu, tearoff=0)

        # Populate the first column (1-15)
        for day in range(1, 16):
            day_menu_column1.add_command(label=str(day), command=lambda d=day: self.day_var.set(d))

        # Populate the second column (16-31)
        for day in range(16, 32):
            day_menu_column2.add_command(label=str(day), command=lambda d=day: self.day_var.set(d))

        # We now add the sub-menus to the main menus
        day_menu.add_cascade(label="1-15", menu=day_menu_column1)
        day_menu.add_cascade(label="16-31", menu=day_menu_column2)


        # Styling the menu's above
        year_menu.config(fg="blue",bg="orange",font=("Helvatica", 12))
        month_menu.config(fg="blue",bg="orange",font=("Helvatica", 12))
        day_menu_button.config(fg="blue",bg="orange",font=("Helvatica", 12))
        day_menu_column1.config(fg="blue",bg="lightgray",font=("Helvatica", 12))
        day_menu_column2.config(fg="blue",bg="lightgray",font=("Helvatica", 12))

        # Output file specification
        output_file_label = tk.Label(scrollable_frame.scrollable_frame, text="Specify Output File:")
        output_file_label.pack(pady=5)

        # Browse output file
        browse_output_button = tk.Button(scrollable_frame.scrollable_frame, text="Browse", command=lambda: browse_output_file(self))
        browse_output_button.pack(pady=5)

        self.output_file_entry = tk.Entry(scrollable_frame.scrollable_frame, width=40)
        self.output_file_entry.pack(pady=5)

        # Button to browse for input file
        def browse_input_file(self):
            filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            self.input_file_entry.delete(0, tk.END) # Clear current entry
            self.input_file_entry.insert(0, filename) # Insert selected file path


        def browse_output_file(self):
            filename = filedialog.asksaveasfilename(defaultextension=".xlsx",filetypes=[("Excel Files", "*.xlsx;*.xls")])
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, filename)
        

        # Add event handlers for button clicks, etc.
        def extract_button_click(self):
            input_file = self.input_file_entry.get()
            output_file = self.output_file_entry.get()
            columns_to_extract = columns_extracted() # Based on the function output
            
            if not input_file:
                messagebox.showwarning("Input Error", "Please select an input file.")
                return
            try:
                selected_date_str = f"{self.year_var.get()}-{self.month_var.get()}-{self.day_var.get()}"
                date_filter = datetime.strptime(selected_date_str, "%Y-%m-%d")

                extract_data(input_file, output_file, columns_to_extract, self.customer_var.get(), date_filter)

            except ValueError as ve:
                messagebox.showwarning("Input Error", str(ve))

        # Button to extract data
        extract_button = tk.Button(scrollable_frame.scrollable_frame, text="Extract Data", command=lambda: extract_button_click())
        extract_button.pack(pady=10)


# Start the GUI event loop
if __name__ == "__main__":
    app = App()
    app.mainloop()
