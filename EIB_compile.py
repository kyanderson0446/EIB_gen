from glob import glob
import openpyxl as oxl
import pandas as pd
import xlwings
import xlwings as xw
import sys
import time
import win32com.client
import datetime

app = xw.App(add_book=False)
xlwings.App.display_alerts = False
main_eib = xw.Book(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\WD_upload_budget_main.xlsx")
path = fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\2028\*.xlsx"
x = 6


for file in glob(path):
    budget_wb = xw.Book(file, update_links=False)
    upload_page = budget_wb.sheets("Budget Lines Data")

    upload_data = upload_page.range("B6:M1241")
    time.sleep(1)

    main_eib.sheets["Budget Lines Data"].range(f"B{x}:M{x}").value = upload_data.value
    time.sleep(1)

    x += 1236

    budget_wb.close()
print("STOP \n " * 3 )

main_eib.save(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\WD_upload_budget_batch_Q1-28.xlsx")


