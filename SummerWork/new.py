# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import xlrd
import xlsxwriter
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
#----------------------------------------------------------

#data = pd.Excelfile(SummerWork.xlsx)
dfs = pd.read_excel(r'C:\Users\kotharan\Documents\GitHub\qualtrics_data_py\SummerWork\All_Reports\Adrian_SD_61_2017-2018_Outdoor_School_Report.xlsx', sheet_name=None)


book = xlrd.open_workbook(r'C:\Users\kotharan\Documents\GitHub\qualtrics_data_py\SummerWork\All_Reports\Adrian_SD_61_2017-2018_Outdoor_School_Report.xlsx')
worksheet = book.sheet_by_name('Adrian School District #61')
num_rows = worksheet.nrows # shows the number of rows in a sheets

# print (num_rows) # to print number of rows

# print (book.nsheets)  ---------------- to print number of sheets
#print ("Sheet Name = " + str(book.sheet_names())) # ---------------- to print all sheet names

first_sheet = book.sheet_by_index(0) # ---------get the first worksheet
print ("Sheet Name:",first_sheet.name)
 # read a row
#print (first_sheet.row_values(0))
sheet1_Data = [first_sheet.cell(1,1),first_sheet.cell(1,3),first_sheet.cell(8,2),first_sheet.cell(8,4),first_sheet.cell(9,2),first_sheet.cell(9,4),first_sheet.cell(10,2),first_sheet.cell(10,4)]
print ("School Name: " + sheet1_Data[0].value)
print ("District Name: " + sheet1_Data[1].value)

# if first_sheet.sheet1_Data(1,0).value == xlrd.empty_sheet1_Data.value:  # TO CHECK FOR EMPTY CELLS
#     print ("Its an empty sheet1_Data")
# else:
#     print("Not empty")

#===========CREATING A WORKBOOK TO STORE======================#

workbook = xlsxwriter.Workbook('OUTPUT.xlsx') #Creating new workbook to write into

#This sheet will store the district name , schoolname and option num. 2 data of the original doc

worksheet1 = workbook.add_worksheet('Sch & Dis Contact Info') #Creating new worksheet to write into
worksheet1.set_column(0,200,10)   # Columns [Start-column,End-column,Width] width set to 30.


worksheet1.write('A1', "School Name :")      #Print Scool Name
worksheet1.write('B1', "District Name :" )
worksheet1.write('A2', sheet1_Data[0].value )        #Taking values from our document spreedsheet and printing in the newly created xslx file
worksheet1.write('B2', sheet1_Data[1].value )

# Taking Cell values from the document and printing in the created OUTPUT.xlsx file
worksheet1.write('C1', "School contact name: " )
worksheet1.write('C2', sheet1_Data[2].value )

worksheet1.write('D1', "District/ESD contact name" )
worksheet1.write('D2', sheet1_Data[3].value )

worksheet1.write('E1', "School contact email" )
worksheet1.write('E2', sheet1_Data[4].value )

worksheet1.write('F1', "District/ESD contact email" )
worksheet1.write('F2', sheet1_Data[5].value )

worksheet1.write('G1', "School contact phone" )
worksheet1.write('G2', sheet1_Data[6].value )

worksheet1.write('H1', "District/ESD contact phone" )
worksheet1.write('H2', sheet1_Data[7].value )


#This sheet will store data from option 3 to 16
sheet2_Data = [first_sheet.cell(1,1)] # Redefining the same variable cell to use it for another worksheet

worksheet2 = workbook.add_worksheet('Other Info') #Creating the second worksheet to write into
#worksheet2.set_column(0, 1, 30)   # Columns [Start-column,End-column,Width] width set to 30.



worksheet2.write('A2', "School Name")      #Print Scool Name
worksheet2.write('B2', "District Name" )
worksheet2.write('A3', sheet1_Data[0].value )        #Taking values from our document spreedsheet and printing in the newly created xslx file
worksheet2.write('B3', sheet1_Data[1].value )

cell_format = workbook.add_format({'color':'red'})


worksheet2.write('C1', "Topic/ Concepts :                                                      " ,cell_format)

worksheet2.write('C2', "Soil,Water, Plants, & Animals" )
worksheet2.write('D2', "Role of timber, agriculture, and other natural resources in the economy of this state" )
worksheet2.write('E2', "The interrelationship of nature, natural resources, economic development and career opportunities in this state" )
worksheet2.write('F2', "The importance of this state’s environment and natural resources" )
worksheet2.write('G2', "The development of students’ leadership, critical thinking and decision-making skills" )



#Continuew from adding values  of the option 6 

workbook.close()
