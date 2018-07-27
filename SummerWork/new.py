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
dfs = pd.read_excel(r'C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork\All_Reports\Adrian_SD_61_2017-2018_Outdoor_School_Report.xlsx', sheet_name=None)

# Book is the name of the workbook of the raw data

book = xlrd.open_workbook(r'C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork\All_Reports\Adrian_SD_61_2017-2018_Outdoor_School_Report.xlsx')
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
sheet2_Data = [first_sheet.cell(17,6),first_sheet.cell(18,6),first_sheet.cell(19,6),first_sheet.cell(20,6),first_sheet.cell(21,6),first_sheet.cell(26,2),first_sheet.cell(27,2),first_sheet.cell(28,2),first_sheet.cell(29,2),first_sheet.cell(30,2),first_sheet.cell(31,2),first_sheet.cell(32,2),first_sheet.cell(26,5),first_sheet.cell(27,5),first_sheet.cell(28,5),first_sheet.cell(29,5),first_sheet.cell(30,5),first_sheet.cell(31,5),first_sheet.cell(34,1),first_sheet.cell(35,1),first_sheet.cell(37,1),first_sheet.cell(38,1),first_sheet.cell(40,1),first_sheet.cell(41,1),first_sheet.cell(43,1),first_sheet.cell(44,1),first_sheet.cell(46,1),first_sheet.cell(47,2),first_sheet.cell(48,2),first_sheet.cell(49,2),first_sheet.cell(50,2),first_sheet.cell(51,2),first_sheet.cell(52,2),first_sheet.cell(53,2),first_sheet.cell(55,3),first_sheet.cell(57,1),first_sheet.cell(58,1),first_sheet.cell(60,1),first_sheet.cell(61,1),first_sheet.cell(63,1),first_sheet.cell(64,1),first_sheet.cell(64,2),first_sheet.cell(65,1),first_sheet.cell(65,2),first_sheet.cell(66,1),first_sheet.cell(66,2),first_sheet.cell(67,1),first_sheet.cell(67,2),first_sheet.cell(68,1),first_sheet.cell(68,2),first_sheet.cell(69,1),first_sheet.cell(69,2),first_sheet.cell(70,1),first_sheet.cell(70,2),first_sheet.cell(71,1),first_sheet.cell(71,2),first_sheet.cell(72,1),first_sheet.cell(72,2),first_sheet.cell(73,1),first_sheet.cell(73,2),first_sheet.cell(74,1),first_sheet.cell(74,2),first_sheet.cell(75,1),first_sheet.cell(75,2),first_sheet.cell(76,1),first_sheet.cell(76,2),first_sheet.cell(77,1),first_sheet.cell(77,2),first_sheet.cell(79,1),first_sheet.cell(80,1),first_sheet.cell(80,2),first_sheet.cell(81,1),first_sheet.cell(81,2),first_sheet.cell(82,1),first_sheet.cell(82,2),first_sheet.cell(83,1),first_sheet.cell(83,2),first_sheet.cell(84,1),first_sheet.cell(84,2),first_sheet.cell(85,1),first_sheet.cell(85,2),first_sheet.cell(86,1),first_sheet.cell(86,2),first_sheet.cell(87,1),first_sheet.cell(87,2),first_sheet.cell(88,1),first_sheet.cell(88,2),first_sheet.cell(89,1),first_sheet.cell(89,2),first_sheet.cell(90,1),first_sheet.cell(90,2),first_sheet.cell(92,1),first_sheet.cell(93,1),first_sheet.cell(95,1),first_sheet.cell(96,1),first_sheet.cell(98,1),first_sheet.cell(99,1)] # Redefining the same variable cell to use it for another worksheet

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

worksheet2.write('C3', sheet2_Data[0].value)
worksheet2.write('D3', sheet2_Data[1].value)
worksheet2.write('E3', sheet2_Data[2].value)
worksheet2.write('F3', sheet2_Data[3].value)
worksheet2.write('G3', sheet2_Data[4].value)

worksheet2.write('H1', "Program Content :                                                      " ,cell_format)

worksheet2.write('H2',"Science" )
worksheet2.write('I2',"Socail Science" )
worksheet2.write('J2',"Food & Agriculture" )
worksheet2.write('K2',"Forestry" )
worksheet2.write('L2',"Sustainability/Enviromental Ed" )
worksheet2.write('M2',"Education Arts" )
worksheet2.write('N2',"Language Arts" )
worksheet2.write('O2',"Math" )
worksheet2.write('P2',"Geography" )
worksheet2.write('Q2',"STEM/STEAM" )
worksheet2.write('R2',"Visual and Performing Arts" )
worksheet2.write('S2',"Physical/Health Ed" )
worksheet2.write('T2',"Other" )


worksheet2.write('H3', sheet2_Data[5].value)
worksheet2.write('I3', sheet2_Data[6].value)
worksheet2.write('J3', sheet2_Data[7].value)
worksheet2.write('K3', sheet2_Data[8].value)
worksheet2.write('L3', sheet2_Data[9].value)
worksheet2.write('M3', sheet2_Data[10].value)
worksheet2.write('N3', sheet2_Data[11].value)
worksheet2.write('O3', sheet2_Data[12].value)
worksheet2.write('P3', sheet2_Data[13].value)
worksheet2.write('Q3', sheet2_Data[14].value)
worksheet2.write('R3', sheet2_Data[15].value)
worksheet2.write('S3', sheet2_Data[16].value)
worksheet2.write('T3', sheet2_Data[17].value)


worksheet2.write('U1', sheet2_Data[18].value, cell_format) # This is the question for option 5 from the main spreadsheet
worksheet2.write('U3', sheet2_Data[19].value) # This is the answer for option 5 from the main spreadsheet

worksheet2.write('V1', sheet2_Data[20].value,cell_format) # This is the question for option 6 from the main spreadsheet
worksheet2.write('V3', sheet2_Data[21].value) # This is the answer for option 6 from the main spreadsheet

worksheet2.write('W1', sheet2_Data[22].value,cell_format) # This is the question for option 7 from the main spreadsheet
worksheet2.write('W3', sheet2_Data[23].value) # This is the answer for option 7 from the main spreadsheet

worksheet2.write('X1', sheet2_Data[24].value,cell_format) # This is the question for option 8 from the main spreadsheet
worksheet2.write('X3', sheet2_Data[25].value) # This is the answer for option 8 from the main spreadsheet

worksheet2.write('Y1', sheet2_Data[26].value,cell_format) # This is the question for option 9 from the main spreadsheet
worksheet2.write('Y2', "Project based Learning") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('Y3', sheet2_Data[27].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('Z2', "Cooperative learning stategies") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('Z3', sheet2_Data[28].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AA2', "Service Learning") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AA3', sheet2_Data[29].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AB2', "Interdisciplinary instruction") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AB3', sheet2_Data[30].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AC2', "Inquiry-based instruction") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AC3', sheet2_Data[31].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AD2', "Social Emotional learning") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AD3', sheet2_Data[32].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AE2', "Socio scientific issues") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AE3', sheet2_Data[33].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AF2', "Other (list)") # This is the subpart for option 9 from the main spreadsheet
worksheet2.write('AF3', sheet2_Data[34].value) # This is the answer for option 9 from the main spreadsheet

worksheet2.write('AG1',sheet2_Data[35].value,cell_format ) # This is the question for option 10 from the main spreadsheet
worksheet2.write('AG3', sheet2_Data[36].value) # This is the answer for option 10 from the main spreadsheet

worksheet2.write('AH1',sheet2_Data[37].value,cell_format ) # This is the question for option 11 from the main spreadsheet
worksheet2.write('AH3', sheet2_Data[38].value) # This is the answer for option 11 from the main spreadsheet

worksheet2.write('AI1', sheet2_Data[39].value,cell_format) # This is the question for option 12 from the main spreadsheet
worksheet2.write('AI2', sheet2_Data[40].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AI3', sheet2_Data[41].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AJ2', sheet2_Data[42].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AJ3', sheet2_Data[43].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AK2', sheet2_Data[44].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AK3', sheet2_Data[45].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AL2', sheet2_Data[46].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AL3', sheet2_Data[47].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AM2', sheet2_Data[48].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AM3', sheet2_Data[49].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AN2', sheet2_Data[50].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AN3', sheet2_Data[51].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AO2', sheet2_Data[52].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AO3', sheet2_Data[53].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AP2', sheet2_Data[54].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AP3', sheet2_Data[55].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AQ2', sheet2_Data[56].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AQ3', sheet2_Data[57].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AR2', sheet2_Data[58].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AR3', sheet2_Data[59].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AS2', sheet2_Data[60].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AS3', sheet2_Data[61].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AT2', sheet2_Data[62].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AT3', sheet2_Data[63].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AU2', sheet2_Data[64].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AU3', sheet2_Data[65].value) # This is the answer for option 12 from the main spreadsheet

worksheet2.write('AV2', sheet2_Data[66].value) # This is the subpart for option 12 from the main spreadsheet
worksheet2.write('AV3', sheet2_Data[67].value) # This is the answer for option 12 from the main spreadsheet

#################################################33333
#######################################################

worksheet2.write('AW1', sheet2_Data[68].value,cell_format) # This is the question for option 13 from the main spreadsheet
worksheet2.write('AW2', sheet2_Data[69].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('AW3', sheet2_Data[70].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('AX2', sheet2_Data[71].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('AX3', sheet2_Data[72].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('AY2', sheet2_Data[73].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('AY3', sheet2_Data[74].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('AZ2', sheet2_Data[75].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('AZ3', sheet2_Data[76].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BA2', sheet2_Data[77].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BA3', sheet2_Data[78].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BB2', sheet2_Data[79].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BB3', sheet2_Data[80].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BC2', sheet2_Data[81].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BC3', sheet2_Data[82].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BD2', sheet2_Data[83].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BD3', sheet2_Data[84].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BE2', sheet2_Data[85].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BE3', sheet2_Data[86].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BF2', sheet2_Data[87].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BF3', sheet2_Data[88].value) # This is the answer for option 13 from the main spreadsheet

worksheet2.write('BG2', sheet2_Data[89].value) # This is the subpart for option 13 from the main spreadsheet
worksheet2.write('BG3', sheet2_Data[90].value) # This is the answer for option 13 from the main spreadsheet


worksheet2.write('BH1', sheet2_Data[91].value,cell_format) # This is the question for option 14 from the main spreadsheet
worksheet2.write('BH3', sheet2_Data[92].value) # This is the answer for option 14 from the main spreadsheet

worksheet2.write('BI1', sheet2_Data[93].value,cell_format) # This is the question for option 15 from the main spreadsheet
worksheet2.write('BI3', sheet2_Data[94].value) # This is the answer for option 15 from the main spreadsheet

worksheet2.write('BJ1', sheet2_Data[95].value,cell_format) # This is the question for option 16 from the main spreadsheet
worksheet2.write('BJ3', sheet2_Data[96].value) # This is the answer for option 16 from the main spreadsheet


workbook.close()
