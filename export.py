import csv
from os import unlink
import pylightxl as xl
import pyzipper
from tkinter import Tk, filedialog
from pathlib import Path
from mypasswords import passwords

root = Tk()
root.withdraw()

BIG_ZIP_PATH = Path("my_zip.zip")

MY_EXCEL_FILE = Path(filedialog.askopenfilename(title="Select the excel file with student names in column A"))

db = xl.readxl(fn=MY_EXCEL_FILE, ws=("Speech"))

row_num = 1

big_zip = pyzipper.ZipFile(BIG_ZIP_PATH, mode="w")

for row in db.ws(ws="Speech").rows:
    if row_num == 1:
        header = row
    else:
        student = row[0]
        if len(student) < 2:
            continue
        csv_filename = student + ".csv"
        with open(csv_filename, "w", encoding="utf-16le", newline="\n") as f:
            writer = csv.writer(f)
            writer.writerow(header)
            writer.writerow(row)
        password = passwords[student]
        zip_name = student + ".zip"
        with pyzipper.AESZipFile(zip_name, "w", compression=pyzipper.ZIP_LZMA,
                                encryption=pyzipper.WZ_AES) as myzip:
            myzip.setpassword(password)
            myzip.write(csv_filename)
        big_zip.write(zip_name)
        unlink(zip_name)
        unlink(csv_filename)
    row_num += 1
big_zip.close()
