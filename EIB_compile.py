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


def open_workbook(file_path):
    """Open an Excel workbook and return the Workbook object."""
    return xw.Book(file_path)


def process_facility_files(year_folder, year_input):
    """Process all individual facility files in a given year folder."""
    for filename in os.listdir(year_folder):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(year_folder, filename)
            wb = open_workbook(file_path)

            try:
                sheet = wb.sheets["Budget Lines Data"]
                sheet.range("F6:F1241").value = year_input
                wb.save()
                print(f"Changes saved in {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {e}")
            finally:
                wb.close()

    print("Processing complete.")


def compile_facility_files():
    """Compile data from individual facility files into one main file."""
    app = xw.App(add_book=False)
    xlwings.App.display_alerts = False
    main_eib = open_workbook(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\WD_upload_budget_main.xlsx")
    path = fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\Ross\2024\*.xlsx"
    x = 6

    for file in glob(path):
        print(file)
        budget_wb = open_workbook(file, update_links=False)
        upload_page = budget_wb.sheets("Budget Lines Data")

        upload_data = upload_page.range("B6:M1241")
        time.sleep(1)

        main_eib.sheets["Budget Lines Data"].range(f"B{x}:M{x}").value = upload_data.value
        time.sleep(1)

        x += 1236

        budget_wb.close()

    main_eib.save(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\Ross_compile_12.18.1_2024.xlsx")


def main():
    """Main function to execute the entire process."""
    year_input = int(input("What year?: "))
    year_folder = fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\{year_input}"

    # Process individual facility files
    process_facility_files(year_folder, year_input)

    # Compile data into a main file
    compile_facility_files()


if __name__ == "__main__":
    main()
