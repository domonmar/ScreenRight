import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
from screen_right import format_word_file, load_parameters, save_parameters, load_default_parameters

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

def run_formatting():
    input_path = input_entry.get()
    output_path = output_entry.get()
    if not input_path or not os.path.exists(input_path):
        messagebox.showerror("Error", "Please select a valid input file.")
        return
    try:
        format_word_file(input_path, output_path)
        messagebox.showinfo("Success", f"Formatted file saved as: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def auto_save_params(event=None):
    """Auto-save parameters when they change."""
    params = {}
    for key, entry in param_entries.items():
        params[key] = entry.get()
    try:
        save_parameters(params)
    except Exception as e:
        # Silently handle save errors to avoid interrupting user
        print(f"Failed to auto-save parameters: {str(e)}")

def reset_to_defaults():
    """Reset parameters to defaults."""
    try:
        default_params = load_default_parameters()
        save_parameters(default_params)  # Save defaults as user settings
        # Update GUI
        for key, entry in param_entries.items():
            entry.delete(0, tk.END)
            entry.insert(0, default_params.get(key, ""))
        messagebox.showinfo("Success", "Parameters reset to defaults.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to reset parameters: {str(e)}")

# Load parameters
parameters = load_parameters()

# Create main window
root = tk.Tk()
root.title("ScreenRight Formatter")

# Create notebook for tabs
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Formatting tab
format_frame = ttk.Frame(notebook)
notebook.add(format_frame, text='Formatting')

# Input file selection
tk.Label(format_frame, text="Input File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
input_entry = tk.Entry(format_frame, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(format_frame, text="Browse", command=select_file).grid(row=0, column=2, padx=5, pady=5)

# Output file display
tk.Label(format_frame, text="Output File:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
output_entry = tk.Entry(format_frame, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)

# Run button
tk.Button(format_frame, text="Format Document", command=run_formatting).grid(row=2, column=1, pady=10)

# Parameters tab
params_frame = ttk.Frame(notebook)
notebook.add(params_frame, text='Parameters')

# Create parameter entries
param_entries = {}
row = 0
for key, value in parameters.items():
    tk.Label(params_frame, text=f"{key}:").grid(row=row, column=0, sticky="w", padx=5, pady=2)
    entry = tk.Entry(params_frame, width=30)
    entry.insert(0, value)
    entry.grid(row=row, column=1, padx=5, pady=2)
    # Bind auto-save to focus out and key release events
    entry.bind("<FocusOut>", auto_save_params)
    entry.bind("<KeyRelease>", auto_save_params)
    param_entries[key] = entry
    row += 1

# Reset to defaults button
tk.Button(params_frame, text="Reset to Defaults", command=reset_to_defaults).grid(row=row, column=1, pady=10)

root.mainloop()