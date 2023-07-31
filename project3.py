import openpyxl
import tkinter as tk

def add_data_to_excel(roll_number, name):
    # Open the Excel file or create a new one if it doesn't exist
    try:
        workbook = openpyxl.load_workbook('data.xlsx')
    except FileNotFoundError:
        workbook = openpyxl.Workbook()

    # Select the active sheet (default: first sheet)
    sheet = workbook.active

    # Append the data to the Excel sheet
    row = [roll_number, name]
    sheet.append(row)

    # Save the changes to the Excel file
    workbook.save('data.xlsx')

def on_submit():
    roll_number = roll_entry.get()
    name = name_entry.get()

    try:
        add_data_to_excel(roll_number, name)
        result_label.config(text="Data successfully stored in Excel!", fg="green")
    except Exception as e:
        result_label.config(text=f"Error occurred: {e}", fg="red")

# Create the tkinter window
root = tk.Tk()
root.title("Data Entry")

# Labels and Entry widgets for roll number and name
roll_label = tk.Label(root, text="Roll Number:")
roll_label.pack()
roll_entry = tk.Entry(root)
roll_entry.pack()

name_label = tk.Label(root, text="Name:")
name_label.pack()
name_entry = tk.Entry(root)
name_entry.pack()

submit_button = tk.Button(root, text="Submit", command=on_submit)
submit_button.pack()

result_label = tk.Label(root, text="", fg="green")
result_label.pack()

# Run the tkinter main loop
root.mainloop()
