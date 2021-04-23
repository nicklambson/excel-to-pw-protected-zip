import csv
from os import unlink
from pathlib import Path
from tkinter import Tk, filedialog, simpledialog, StringVar, Radiobutton, X, Button

import pylightxl as xl
import pyzipper

from mypasswords import PASSWORDS

root = Tk()
root.withdraw()

MY_EXCEL_FILE = Path(filedialog.askopenfilename(title="Select the excel file with student names in column A",
                                                filetypes=[("Excel Files", "*.xlsx")]))

db = xl.readxl(fn=MY_EXCEL_FILE)
ws_names_str = ", ".join(db.ws_names)

selected_sheet = StringVar(root, "Sheet1")

root.deiconify()


for sheet_name in db.ws_names:
    Radiobutton(root, text = sheet_name, variable = selected_sheet,
                value = sheet_name, indicator = 0,
                background = "light blue").pack(fill = X, ipady = 5)

b1 = Button(root, text="OK", command=root.destroy)
b1.pack()
# TODO: fix title
# root.title("Please select the Excel Sheet you'd like to read.")
root.mainloop()

my_ws = db.ws(ws=selected_sheet.get())

BIG_ZIP_PATH = Path(filedialog.asksaveasfilename(title="Name the zip file to save."))

row_num = 1

big_zip = pyzipper.ZipFile(BIG_ZIP_PATH, mode="w")

for row in my_ws.rows:
    if row_num == 1:
        header = row
    else:
        student = row[0].strip()
        if len(student) < 2:
            continue
        csv_name = student + ".csv"
        csv_path = Path.home() / csv_name
        with open(csv_path, "w", encoding="utf-16le", newline="\n") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerow(row)
        password = PASSWORDS[student]
        zip_name = student + ".zip"
        zip_path = Path.home() / zip_name
        with pyzipper.AESZipFile(zip_path, "w", compression=pyzipper.ZIP_LZMA,
                                encryption=pyzipper.WZ_AES) as myzip:
            myzip.setpassword(password)
            myzip.write(csv_path, csv_name)
        big_zip.write(zip_path, zip_name)
        unlink(zip_path)
        unlink(csv_path)
    row_num += 1
big_zip.close()
