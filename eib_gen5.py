from glob import glob
import openpyxl as oxl
import pandas as pd
import xlwings
import xlwings as xw
import sys
import time
import win32com.client
import datetime
import os

# app = xw.App(add_book=False)
# xlwings.App.display_alerts = False

# Accessing Files and Pdrive
print("Access Folder")
folder = str(input("Enter \"YYYY Quarter #\": "))
quarter_entry = folder.split(" ")
quarter_entry = quarter_entry[1]

year_entry = folder.split(" ")
year_entry = year_entry[0]
year_entry = int(year_entry)

# get the year and once each year forecast is complete, loop through current year+x

eib_max = 65
print("*" * 60)

# path = fr"P:\PACS\Finance\Budgets\{folder}\Received\Final Q1\*.xlsx"
df_fac = pd.read_csv("Facility_list.csv")
df_fac = df_fac['Facility']
facilities = list(df_fac)



EIB_wb = xw.Book(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney PACS_Import_Budget-test-main 5 year - Copy.xlsm")

# start_date_key = {'Q1': f'1/1/{year_entry}', 'Q2': f'4/1/{year_entry}',
#                   'Q3': f'7/1/{year_entry}', 'Q4': f'10/1/{year_entry}'}

# Begin For Loop of Opening Budgets and Pasting Values Into EIB
x = 1
current_year = 0

bhag = 1
while bhag < 6:
    for facility in facilities:
        # if os.path.exists(
        #         fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\{year_entry}\{year_entry} {facility} EIB_budget.xlsx"):
        #     print(fr"{year_entry} {facility} exists")
        #     continue
        print(fr"Creating file for {year_entry} {facility}")
        EIB_wb.sheets["WD Upload"].range(f"A1").value = facility
        time.sleep(1)
        current_quarter_date = fr"1/1/{year_entry}"
        time.sleep(1)
        EIB_wb.sheets["WD Upload"].range(f"C2").value = current_quarter_date
        #year start date
        EIB_wb.sheets["WD Upload"].range(f"C3").value = year_entry
        time.sleep(5)
        # Run Macros
        macro2 = EIB_wb.macro("Module2.Clear_all")
        macro2()
        time.sleep(3)
        macro1 = EIB_wb.macro("Module1.Run_refresh")
        macro1()
        time.sleep(10)
        macro1 = EIB_wb.macro("Module1.GL_code")
        macro1()
        time.sleep(6)
        macro1 = EIB_wb.macro("Module1.Credit_amount")
        macro1()
        time.sleep(1)
        macro1 = EIB_wb.macro("Module1.Debit_amount")
        macro1()
        time.sleep(1)
        macro1 = EIB_wb.macro("Module1.Month_drop")
        macro1()
        time.sleep(2)

        try:
            os.makedirs(
                fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\{year_entry}")
        except:
            pass

        EIB_wb.save(fr"C:\Users\kyle.anderson\OneDrive - PACS\backup\Documents\Courtney\{year_entry}\{year_entry} {facility} EIB_budget.xlsx")   # change quarter
        # EIB_wb.close()
    year_entry += 1
    bhag += 1

print()
print(time.process_time(), "minutes")

