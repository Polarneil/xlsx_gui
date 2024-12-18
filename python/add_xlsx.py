import openpyxl
from openpyxl import Workbook
import tkinter as tk
from tkinter import ttk, messagebox

# Function to add a record to Excel
def add_record_to_excel(file_path, record_name, project_manager, days_to_complete, importance):
    new_row = [record_name, project_manager, days_to_complete, importance]
    try:
        # Load workbook if it exists, otherwise create a new one
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = Workbook()
            sheet = workbook.active
            # Add header row
            sheet.append(["Record Name", "Project Manager", "Days to Complete", "Importance"])
        
        # Append the new row
        sheet.append(new_row)
        # Save the file
        workbook.save(file_path)
        messagebox.showinfo("Success", f"Record added successfully to '{file_path}'")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# GUI Function
def open_gui():
    def check_importance():
        selected = [var_low.get(), var_medium.get(), var_high.get()]
        if sum(selected) > 2:  # If more than 2 checkboxes are selected
            messagebox.showwarning("Selection Error", "You can select up to 2 options only!")
            return False
        return True

    # submit function - handle data input
    def submit():
        record_name = record_name_entry.get()
        project_manager = project_manager_entry.get()
        days_to_complete = int(days_combo.get())

        # Collect Importance Selections
        importance_selections = []
        if var_low.get():
            importance_selections.append("Low")
        if var_medium.get():
            importance_selections.append("Medium")
        if var_high.get():
            importance_selections.append("High")

        if len(importance_selections) > 2:
            messagebox.showwarning("Input Error", "You can select up to 2 importance levels!")
            return

        importance = ", ".join(importance_selections)

        if not record_name or not project_manager or not days_to_complete or not importance:
            messagebox.showwarning("Input Error", "All fields must be filled!")
            return
        
        add_record_to_excel("xlsx/add_record.xlsx", record_name, project_manager, days_to_complete, importance)
        # clear fields after submit
        record_name_entry.delete(0, tk.END)
        project_manager_entry.delete(0, tk.END)
        days_combo.set("")
        var_low.set(0)
        var_medium.set(0)
        var_high.set(0)

    # Main window
    root = tk.Tk()
    root.title("Simple Data Entry Form")
    root.geometry("650x350")
    root.configure(bg="#1E3D59")  # Dark blue background

    # Styling configuration
    label_font = ("Arial", 12, "bold")
    input_font = ("Arial", 12)
    fg_color = "white"
    btn_bg = "#4CAF50"

    # Grid Layout Configuration
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=3)

    # Title Label
    title_label = tk.Label(root, text="Please fill out the following fields:", font=("Arial", 14, "bold"), bg="#1E3D59", fg=fg_color)
    title_label.grid(row=0, column=0, columnspan=2, pady=(10, 20), sticky="w")

    # Name Input
    tk.Label(root, text="Name", font=label_font, bg="#1E3D59", fg=fg_color).grid(row=1, column=0, padx=20, pady=5, sticky="w")
    record_name_entry = tk.Entry(root, font=input_font, width=30)
    record_name_entry.grid(row=1, column=1, padx=20, pady=5, sticky="w")

    # Project Manager Input
    tk.Label(root, text="Project Manager", font=label_font, bg="#1E3D59", fg=fg_color).grid(row=2, column=0, padx=20, pady=5, sticky="w")
    project_manager_entry = tk.Entry(root, font=input_font, width=30)
    project_manager_entry.grid(row=2, column=1, padx=20, pady=5, sticky="w")

    # Days to Complete Dropdown
    tk.Label(root, text="Days to Complete", font=label_font, bg="#1E3D59", fg=fg_color).grid(row=3, column=0, padx=20, pady=5, sticky="w")
    days_combo = ttk.Combobox(root, values=[30, 60, 90], state="readonly", font=input_font, width=28)
    days_combo.grid(row=3, column=1, padx=20, pady=5, sticky="w")

    # Importance Checkboxes
    tk.Label(root, text="Importance (select up to 2)", font=label_font, bg="#1E3D59", fg=fg_color).grid(row=4, column=0, padx=20, pady=5, sticky="nw")

    var_low = tk.IntVar()
    var_medium = tk.IntVar()
    var_high = tk.IntVar()

    tk.Checkbutton(root, text="Low", variable=var_low, onvalue=1, offvalue=0, bg="#1E3D59", fg=fg_color, font=input_font, command=check_importance).grid(row=4, column=1, sticky="w", padx=20)
    tk.Checkbutton(root, text="Medium", variable=var_medium, onvalue=1, offvalue=0, bg="#1E3D59", fg=fg_color, font=input_font, command=check_importance).grid(row=4, column=1, sticky="w", padx=100)
    tk.Checkbutton(root, text="High", variable=var_high, onvalue=1, offvalue=0, bg="#1E3D59", fg=fg_color, font=input_font, command=check_importance).grid(row=4, column=1, sticky="w", padx=180)

    # Submit and Exit Buttons
    button_frame = tk.Frame(root, bg="#1E3D59")
    button_frame.grid(row=5, column=0, columnspan=2, padx=20, pady=20, sticky="w")

    submit_btn = tk.Button(button_frame, text="Submit", command=submit, font=label_font, bg=btn_bg, fg="#1E3D59", padx=10, pady=5)
    submit_btn.pack(side="left", padx=5)

    exit_btn = tk.Button(button_frame, text="Exit", command=root.quit, font=label_font, bg="#E74C3C", fg="#1E3D59", padx=10, pady=5)
    exit_btn.pack(side="left", padx=5)

    root.mainloop()

open_gui()
