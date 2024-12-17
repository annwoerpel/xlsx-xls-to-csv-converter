import os
import tkinter as tk
import pandas as pd
from tkinter import filedialog

# Main window
root = tk.Tk()
root.title("Excel zu CSV Konvertierung")
root.geometry("600x400")
root.config(bg="#f0f0f0")

# Textbox to see what should be exported
text = tk.Text(root, height=20, width=50)
text.pack()
text.insert(tk.END, "Ausgewählte Dateien:\n")

# Convert Excel to CSV and save it
def excel_to_csv():
    input_file_paths = filedialog.askopenfilenames(title = "Dateien auswählen", multiple=True)

    for file_path in input_file_paths:
        name = file_path.split("/")[-1]
        text.insert(tk.END, str(name) + "\n")

    for file_path in input_file_paths:
        file_name, file_extension = os.path.splitext(file_path)
    
        if file_extension == '.xlsx':
            df = pd.read_excel(file_path, engine='openpyxl')
            output_file_name = file_path.split("/")[-1][:-5]
        elif file_extension == '.xls':
            df = pd.read_excel(file_path)
            output_file_name = file_path.split("/")[-1][:-4]
        else:
            raise Exception("File not supported")

        output_file_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=output_file_name)
        df.to_csv(output_file_path, encoding='utf-8-sig', index=False, sep=";")
        
    text.insert(tk.END, "\nDie Dateien wurden konvertiert!")


# Button
button = tk.Button(root, text="Wähle Dateien zum Konvertieren", command=excel_to_csv, font=("Calibri", 14), bg="#507289", fg="#ffffff")
button.pack(pady=20)

# add file dialog
filedialog = tk.filedialog

# run GUI
root.mainloop()

