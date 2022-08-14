import utils

from openpyxl import load_workbook

wb = load_workbook(filename = 'HM_Salary_2111.xlsx')

ws_basistable = wb["1.기본정보TABLE"]

print(ws_basistable)