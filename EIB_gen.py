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

# Accessing Files and Pdrive
print("Access Folder")
folder = str(input("Enter \"YYYY Quarter #\": "))
eib_max = 65
print("*" * 60)

path = fr"P:\PACS\Finance\Budgets\{folder}\Received\Final Q1\*.xlsx"
EIB_wb = xw.Book(fr"C:\Users\kyle.anderson\Documents\Courtney\Courtney PACS_Import_Budget-test.xlsm")

# Begin For Loop of Opening Budgets and Pasting Values Into EIB
x = 1

for file in glob(path):
    budget_wb = xw.Book(file, update_links=False)
    upload_page = budget_wb.sheets("DW Upload")
    facility_info = budget_wb.sheets("FACILITY INFO")

    upload_name = facility_info.range("B7").value
    print(upload_name)
    budget_data = upload_page.range("A1:AB374")

    EIB_wb.sheets["WD Upload"].range(f"A1:AB374").value = budget_data.value

# Run Macros
    macro2 = EIB_wb.macro("Module2.Clear_all")
    macro2()
    time.sleep(1)
    macro1 = EIB_wb.macro("Module1.GL_reformat")
    macro1()
    time.sleep(1)
    macro1 = EIB_wb.macro("Module1.GL_code")
    macro1()
    time.sleep(1)
    macro1 = EIB_wb.macro("Module1.Credit_amount")
    macro1()
    time.sleep(1)
    macro1 = EIB_wb.macro("Module1.Debit_amount")
    macro1()
    time.sleep(1)
    macro1 = EIB_wb.macro("Module1.Month_drop")
    macro1()
    time.sleep(2)

    EIB_wb.save(fr"C:\Users\kyle.anderson\Documents\Courtney\{upload_name} EIB_budget.xlsx")
    budget_wb.close()

    time.sleep(1)


print()
print(time.process_time(), "minutes")

