# Overview
A teacher needs to send confidential information from an excel sheet to each student. Instead of sending information one-by-one, this script writes a unique password protected zip for each student. The password protected zips can be zipped up and sent as one to the whole class. The students can then use a pre-determined password to open the zip with their name on it.

# Prerequisites
We need a lightweight excel reader. pylightxl does a great job. Install it with the pip command below.

`pip install pylightxl`

We also need a zip utility with encryption capabilities. pyzipper mimics the standard library's zipfile, but enables creating zips with passwords. Install it with the pip command below.

`pip install pyzipper`

# Assumptions

 - The user should have an excel sheet with one header row, and one row per student.
 - Student names should be in column A.
 - The sheet with information to be exported should be called "Sheet1" (Excel default).
 - The user has collected the passwords of each student and imported them into the file caled `passwords.py`.

# Operation
1. The user should run the script `export.py`.
2. The user will be prompted to select an excel file to process.
3. The script will read the sheet and write each row to a csv file named with the student's name.
4. The script will zip the csv file using the student's password.
5. The script will zip all the zips into a single zip called my_zip.zip
6. The script will remove all the previously created csv and zip files for cleanup.
7. The user can send the my_zip.zip file to a group of students and each student can open their own file with their own password.

# Notes
This software is provided as-is and no guarantees are made as to its usage.  
TODO: Add "Save as" file functionality for the final name of the big zip.  
TODO: Comment parts of the script

