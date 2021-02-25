import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
book = load_workbook(file_path)
sheet = book.active

for row in sheet:
    for cell in row:
        if cell.value == "Nooit":
            cell.value = 0
        elif cell.value == "Zelden":
            cell.value = 1
        elif cell.value == "Soms":
            cell.value = 2
        elif cell.value == "Vaak":
            cell.value = 3
        elif cell.value == "Zeer vaak":
            cell.value = 4

book.save(file_path)





