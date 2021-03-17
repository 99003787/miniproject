import pandas as pd
import openpyxl
import xlsxwriter
import xlrd
# book = pd.read_excel(r'exam.xlsx')
# print(book)
f = pd.ExcelFile(r'student.xlsx')
sheet1 = f.parse(sheet_name="Sheet1")
all_df = pd.DataFrame()
sheet0 = pd.read_excel(r'student.xlsx', sheet_name="Sheet1")
all_df = all_df.append(sheet0)
print(all_df)
# writer = pd.ExcelWriter('pandasEx.xlsx', engine ='xlsxwriter')
# all_df.to_excel(writer, sheet_name='Sheet5')
#
# # Close the Pandas Excel writer
# # object and output the Excel file.
# writer.save()sheet1 = pd.read_excel(r'student.xlsx', sheet_name="Sheet2")
sheet2 = pd.read_excel(r'student.xlsx', sheet_name="Sheet3")
sheet3 = pd.read_excel(r'student.xlsx', sheet_name="Sheet4")
sheet4 = pd.read_excel(r'student.xlsx', sheet_name="Sheet5")
all_df = pd.merge(all_df, sheet1, how='left')
# print(all_df)
all_df = pd.merge(all_df, sheet2, how='left')
# print(all_df)
all_df = pd.merge(all_df, sheet3, how='left')
# print(all_df)
all_df = pd.merge(all_df, sheet4, how='left')
print(all_df)
writer = pd.ExcelWriter('pandasEx.xlsx', engine ='xlsxwriter')
all_df.to_excel(writer, sheet_name='Sheet5')
