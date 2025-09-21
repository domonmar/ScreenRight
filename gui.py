import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
from screen_right import format_word_file

def select_file():
    file_path = filedialog.askopenfilename(
        title="Select a .docx file",
        filetypes=[("Word documents", "*.docx")]
    )
    if file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)
        # Determine output path
        output_path = os.path.splitext(file_path)[0] + "_out.docx"
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_path)
        # Determine param file path (assuming in same dir as script)
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        elif __file__:
            application_path = os.path.dirname(__file__)
        param_file = os.path.join(application_path, "parameters.txt")
        param_entry.delete(0, tk.END)
        param_entry.insert(0, param_file)

def run_formatting():
    input_path = input_entry.get()
    output_path = output_entry.get()
    param_file = param_entry.get()
    if not input_path or not os.path.exists(input_path):
        messagebox.showerror("Error", "Please select a valid input file.")
        return
    if not os.path.exists(param_file):
        messagebox.showerror("Error", f"Parameters file not found: {param_file}")
        return
    try:
        format_word_file(input_path, output_path, param_file)
        messagebox.showinfo("Success", f"Formatted file saved as: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create main window
root = tk.Tk()
root.title("ScreenRight Formatter")

# Input file selection
tk.Label(root, text="Input File:").grid(row=0, column=0, sticky="w")
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1)
tk.Button(root, text="Browse", command=select_file).grid(row=0, column=2)

# Output file display
tk.Label(root, text="Output File:").grid(row=1, column=0, sticky="w")
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1)

# Param file display
tk.Label(root, text="Parameters File:").grid(row=2, column=0, sticky="w")
param_entry = tk.Entry(root, width=50)
param_entry.grid(row=2, column=1)

# Run button
tk.Button(root, text="Format Document", command=run_formatting).grid(row=3, column=1, pady=10)

root.mainloop()