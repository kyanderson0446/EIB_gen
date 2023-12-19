from glob import glob
import os
import openpyxl as oxl
import pandas as pd
import xlwings
import xlwings as xw
import sys
import time
import win32com.client
import datetime

"""
This process takes all individual Facility files and compiles them into 1 file.
The saved file is then uploaded to Workday to upload all facilities at once.
"""


def year_fix():
    year_input = int(input("What year?: "))
    year_folder = fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\{year_input}"
    for filename in os.listdir(year_folder):
        if filename.endswith(".xlsx"):  # Assuming your files have the .xlsx extension
            file_path = os.path.join(year_folder, filename)

            # Open the Excel workbook
            wb = xw.Book(file_path)

            try:
                # Access the "Sheet1" (change the sheet name as needed) and set the value in the specified range
                wb.sheets["Budget Lines Data"].range("F6:F1241").value = year_input

                # Save the changes
                wb.save()
                print(f"Changes saved in {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {e}")
            finally:
                # Close the workbook
                wb.close()

    print("Processing complete.")



################################
app = xw.App(add_book=False)
xlwings.App.display_alerts = False
main_eib = xw.Book(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\WD_upload_budget_main.xlsx")
path = fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\Ross\2024\*.xlsx"
x = 6

################################
for file in glob(path):
    year_fix()
    print(file)
    budget_wb = xw.Book(file, update_links=False)
    upload_page = budget_wb.sheets("Budget Lines Data")

    upload_data = upload_page.range("B6:M1241")
    time.sleep(1)

    main_eib.sheets["Budget Lines Data"].range(f"B{x}:M{x}").value = upload_data.value
    time.sleep(1)

    x += 1236

    budget_wb.close()
print("STOP \n " * 3)
################################
main_eib.save(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\Ross_compile_12.18.1_2024.xlsx")

    
