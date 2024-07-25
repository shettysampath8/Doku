import sys
import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox

def find_replace(excel_loc, textfile, output_folder):
    flag = False # Indicate inputs are not found
    
    # Read the Excel file
    read_file = pd.read_excel(excel_loc)
    input_val = list(read_file['Input']) # Read input column
    output_val = list(read_file['Output']) # Read output column
    
    # Create the new folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Define output text file path
    output_textfile = os.path.join(output_folder, 'Results.txt')
    
    with open(textfile, 'r') as f: # open file for read
        lines = f.readlines()
    
    with open(output_textfile, 'w') as f: # open file for write
        for line in lines:
            replaced = False
            for i in range(len(input_val)):
                if input_val[i] in line: 
                    st = line.replace(input_val[i], output_val[i])
                    f.write(st)
                    replaced = True
                    flag = True
                    break
            if not replaced:
                f.write(line)
                
    if flag == False:
        messagebox.showinfo("Result", f'The inputs {input_val} are not found in file')
    else:
        messagebox.showinfo("Result", f'Inputs replaced and saved in {output_textfile}')

def browse_excel():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    excel_entry.delete(0, "end")
    excel_entry.insert(0, filename)

def browse_text():
    filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    text_entry.delete(0, "end")
    text_entry.insert(0, filename)

def run():
    excel_loc = excel_entry.get()
    textfile = text_entry.get()
    if not excel_loc or not textfile:
        messagebox.showerror("Error", "Please select both Excel and Text files.")
        return
    output_folder = "output_files"
    find_replace(excel_loc, textfile, output_folder)

# Create the main window
root = Tk()
root.title("Find and Replace Tool")

# Excel file input
excel_label = Label(root, text="Select Excel File:")
excel_label.grid(row=0, column=0, padx=10, pady=10)
excel_entry = Entry(root, width=50)
excel_entry.grid(row=0, column=1, padx=10, pady=10)
excel_button = Button(root, text="Browse", command=browse_excel)
excel_button.grid(row=0, column=2, padx=10, pady=10)

# Text file input
text_label = Label(root, text="Select Text File:")
text_label.grid(row=1, column=0, padx=10, pady=10)
text_entry = Entry(root, width=50)
text_entry.grid(row=1, column=1, padx=10, pady=10)
text_button = Button(root, text="Browse", command=browse_text)
text_button.grid(row=1, column=2, padx=10, pady=10)

# Run button
run_button = Button(root, text="Run", command=run)
run_button.grid(row=2, column=1, pady=20)

# Start the main event loop
root.mainloop()
