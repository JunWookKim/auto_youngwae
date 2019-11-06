import openpyxl
import os

#필요없는 열 제거 -> [소속 이름 대분류 중분류], [소속 이름 대분류 중분류] 리스트 생성 ->

wb = openpyxl.load_workbook("test.xlsx")
sheet = wb.worksheets[0]
#print(sheet.cell(2,2).value) #read value
#sheet.delete_cols(5,2)         #E부터 2개 열 제거
#sheet.delete_cols(6)           #F열 제거


""" 파일 생성 복사~
wb = openpyxl.Workbook()
wb.save('./save_test.xlsx')
os.system("cp save_test.xlsx cp_save_test.xlsx")
os.system("del save_test.xlsx")
"""