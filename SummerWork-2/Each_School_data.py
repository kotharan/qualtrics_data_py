import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter


inputbook = xlrd.open_workbook(r"C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork-2\test.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE======================#

workbook = xlsxwriter.Workbook('.\SummerWork-2\OUTPUT.xlsx') #Creating new workbook to write into
outWorkSheet = workbook.add_worksheet('All Schools Info') #Creating new worksheet to write into

cell_format = workbook.add_format({'color':'red'}) #To do cell text formatting
bold = workbook.add_format({'bold': True})


# For heading

for headcol in range(0,62):
    head_obj = worksheet.cell(1, headcol)  # Get cell object by row, col
    outWorkSheet.write(0,headcol,head_obj.value)


# For printing data

Wrow= 1 # Start of the data
Wcol = 0 # To start from the first column

for row in range(3,14):
    for col in range(0,62):
        cell_obj = worksheet.cell(row, col).value  # Get cell object by row, col
        outWorkSheet.write(Wrow,Wcol,cell_obj)
        Wcol +=1
    Wcol=0
    Wrow+=1
workbook.close()

# For now it is reading till column 62 and then skips the further data of other school and goes to the next line
