import csv, os
import pandas as pd
from pandas import ExcelWriter
import xlrd
import os
from glob import glob
from xlsxwriter.workbook import Workbook
from tabula import convert_into
from xlrd import open_workbook

path3 = 'C://Users//MMCS 9//Desktop//test.xlsx'
path4 = 'C://Users//MMCS 9//Desktop//w.xlsx'

df = pd.read_excel(path3) #Read Excel file as a DataFrame
df['Empty'] = ''
#Display top 5 rows to check if everything looks good
df.head(5)
#To save it back as Excel
df.to_excel(path3) #Write DateFrame back as Excel file

Date = ''
Name = ''
Address = ''
book = open_workbook(path3)
for sheet in book.sheets():
    for rowidx in range(1):
        row = sheet.row(rowidx)
        for colidx, cell in enumerate(row):
            if Date in ('',"Empty"):
                if cell.value in ("Date", "Dt"):
                    Date = cell.value
                else: Date = 'Empty'
            if Name in ('',"Empty"):
                if cell.value in ("Name", "Nme"):
                    Name = cell.value
                else: Name = 'Empty'
            if Address in ('',"Empty"):
                if cell.value in ("Address", "Adrs"):
                    Address = cell.value
                else: Address = 'Empty'

results = pd.read_excel(path3, sheetname="Sheet1")

df = pd.DataFrame({'Date': results[Date],
                   'Name': results[Name],
                   'Address': results[Address]})
writer = ExcelWriter(path4)

df.to_excel(writer, 'Sheet1', index=False)
writer.save()
