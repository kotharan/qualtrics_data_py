import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter

#----------------------------------------------------------



listOfFiles = os.listdir('.\All_Reports')   # To get just the name of the file with .xlsx in All_Reports folder
pattern = "*.xlsx" # To get only files with this extension



#===========CREATING A WORKBOOK TO STORE======================#

workbook = xlsxwriter.Workbook('OUTPUT.xlsx') #Creating new workbook to write into

#This sheet will store the district name , schoolname and option num. 2 data of the original doc

worksheet1 = workbook.add_worksheet('Sch & Dis Contact Info') #Creating new worksheet to write into

cell_format = workbook.add_format({'color':'red'}) #To do cell text formatting
bold = workbook.add_format({'bold': True})

# This funtion stores just the questions and the answers are stored in get_ans funtion
def get_ques(current_sheet):
    main_question=[[worksheet1.write('A1', "School Name :") ,worksheet1.write('B1', "District Name :" ),worksheet1.write('A1', "School Name :"),worksheet1.write('B1', "District Name :" ),worksheet1.write('C1', "School contact name: " ),worksheet1.write('D1', "District/ESD contact name" ),worksheet1.write('E1', "School contact email" ),worksheet1.write('F1', "District/ESD contact email" ),worksheet1.write('G1', "School contact phone" ),worksheet1.write('H1', "District/ESD contact phone"),worksheet1.write('I1', "Topic Content: ",cell_format),worksheet1.write('N1', "Program Content: ",cell_format),worksheet1.write('AE1',current_sheet.cell(46,1).value,cell_format),worksheet1.write('AM1',current_sheet.cell(57,1).value,cell_format),worksheet1.write('AN1',current_sheet.cell(60,1).value,cell_format),worksheet1.write('AO1',current_sheet.cell(63,1).value,cell_format),worksheet1.write('BC1',current_sheet.cell(79,1).value,cell_format),worksheet1.write('BN1',current_sheet.cell(92,1).value,cell_format),worksheet1.write('BO1',current_sheet.cell(95,1).value,cell_format),worksheet1.write('BP1',current_sheet.cell(98,1).value,cell_format)],[current_sheet.cell(34,1),current_sheet.cell(37,1),current_sheet.cell(40,1),current_sheet.cell(43,1)]]

    #THIS LOOP IS TO PRINT ALL THE MAIN QUESTIONS WHICH GET PRINTED ON THE FIRST LINE
    # THE DATA WITH worksheet1.write properties in the main_question[0] array do not need a loop to run
    row1=0
    col1 = 26
    for main_question_arr in range(len(main_question[1])):
        worksheet1.write(row1,col1,main_question[1][main_question_arr].value,cell_format)
        col1 += 1

# Stores the sub questions/sub parts of the main_question
    sub_questions=[worksheet1.write('I2', "Soil,Water, Plants, & Animals",bold ),worksheet1.write('J2', "Role of timber, agriculture, and other natural resources in the economy of this state",bold ),worksheet1.write('K2', "The interrelationship of nature, natural resources, economic development and career opportunities in this state",bold ),worksheet1.write('L2', "The importance of this state’s environment and natural resources",bold ),worksheet1.write('M2', "The development of students’ leadership, critical thinking and decision-making skills",bold),worksheet1.write('N2',"Science",bold ),worksheet1.write('O2',"Social Science",bold ),worksheet1.write('P2',"Food & Agriculture",bold ),worksheet1.write('Q2',"Forestry",bold ),worksheet1.write('R2',"Sustainability/Enviromental Ed",bold ),worksheet1.write('S2',"Education Arts",bold ),worksheet1.write('T2',"Language Arts",bold ),worksheet1.write('U2',"Math",bold ),worksheet1.write('V2',"Geography",bold ),worksheet1.write('W2',"STEM/STEAM",bold ),worksheet1.write('X2',"Visual and Performing Arts",bold ),worksheet1.write('Y2',"Physical/Health Ed",bold ),worksheet1.write('Z2',"Other",bold ),worksheet1.write('AE2', "Project based Learning",bold),worksheet1.write('AF2', "Cooperative learning stategies",bold),worksheet1.write('AG2', "Service Learning",bold),worksheet1.write('AH2', "Interdisciplinary instruction",bold),worksheet1.write('AI2', "Inquiry-based instruction",bold),worksheet1.write('AJ2', "Social Emotional learning",bold),worksheet1.write('AK2', "Socio scientific issues",bold),worksheet1.write('AL2', "Other (list)",bold),worksheet1.write('AO2',current_sheet.cell(64,1).value,bold),worksheet1.write('AP2',current_sheet.cell(65,1).value,bold),worksheet1.write('AQ2',current_sheet.cell(66,1).value,bold),worksheet1.write('AR2',current_sheet.cell(67,1).value,bold),worksheet1.write('AS2',current_sheet.cell(68,1).value,bold),worksheet1.write('AT2',current_sheet.cell(69,1).value,bold),worksheet1.write('AU2',current_sheet.cell(70,1).value,bold),worksheet1.write('AV2',current_sheet.cell(71,1).value,bold),worksheet1.write('AW2',current_sheet.cell(72,1).value,bold),worksheet1.write('AX2',current_sheet.cell(73,1).value,bold),worksheet1.write('AY2',current_sheet.cell(74,1).value,bold),worksheet1.write('AZ2',current_sheet.cell(75,1).value,bold),worksheet1.write('BA2',current_sheet.cell(76,1).value,bold),worksheet1.write('BB2',current_sheet.cell(77,1).value,bold),worksheet1.write('BC2',current_sheet.cell(80,1).value,bold),worksheet1.write('BD2',current_sheet.cell(81,1).value,bold),worksheet1.write('BE2',current_sheet.cell(82,1).value,bold),worksheet1.write('BF2',current_sheet.cell(83,1).value,bold),worksheet1.write('BG2',current_sheet.cell(84,1).value,bold),worksheet1.write('BH2',current_sheet.cell(85,1).value,bold),worksheet1.write('BI2',current_sheet.cell(86,1).value,bold),worksheet1.write('BJ2',current_sheet.cell(87,1).value,bold),worksheet1.write('BK2',current_sheet.cell(88,1).value,bold),worksheet1.write('BL2',current_sheet.cell(89,1).value,bold),worksheet1.write('BM2',current_sheet.cell(90,1).value,bold)]

#Stores all answers executed on and after row 3
def get_ans():
    row =2
    for entry in listOfFiles:
        if fnmatch.fnmatch(entry, pattern): # Stores the name of all .xlsx format in a list so that we can take one file at a time
                name_of_current_workbook = r"C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork\All_Reports" + "\\" + entry # Stored the name of current working workbook
                book = xlrd.open_workbook(name_of_current_workbook) # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files

                for sheet_count in range(len(book.sheet_names())):  # To check if it has more than one sheet, if yes , then go through all of them
                    current_sheet = book.sheet_by_index(sheet_count) # ---------get the first worksheet
                    print ("Sheet Name:",current_sheet.name)

                    get_ques(current_sheet)
                    answers_from_sheet = [current_sheet.cell(1,1),current_sheet.cell(1,3),current_sheet.cell(8,2),current_sheet.cell(8,5),current_sheet.cell(9,2),current_sheet.cell(9,5),current_sheet.cell(10,2),current_sheet.cell(10,5),current_sheet.cell(17,6),current_sheet.cell(18,6),current_sheet.cell(19,6),current_sheet.cell(20,6),current_sheet.cell(21,6),current_sheet.cell(26,2),current_sheet.cell(27,2),current_sheet.cell(28,2),current_sheet.cell(29,2),current_sheet.cell(30,2),current_sheet.cell(31,2),current_sheet.cell(32,2),current_sheet.cell(26,5),current_sheet.cell(27,5),current_sheet.cell(28,5),current_sheet.cell(29,5),current_sheet.cell(30,5),current_sheet.cell(31,5),current_sheet.cell(35,1),current_sheet.cell(38,1),current_sheet.cell(41,1),current_sheet.cell(44,1),current_sheet.cell(47,2),current_sheet.cell(48,2),current_sheet.cell(49,2),current_sheet.cell(50,2),current_sheet.cell(51,2),current_sheet.cell(52,2),current_sheet.cell(53,2),current_sheet.cell(55,3),current_sheet.cell(58,1),current_sheet.cell(61,1),current_sheet.cell(64,2),current_sheet.cell(65,2),current_sheet.cell(66,2),current_sheet.cell(67,2),current_sheet.cell(68,2),current_sheet.cell(69,2),current_sheet.cell(70,2),current_sheet.cell(71,2),current_sheet.cell(72,2),current_sheet.cell(73,2),current_sheet.cell(74,2),current_sheet.cell(75,2),current_sheet.cell(76,2),current_sheet.cell(77,2),current_sheet.cell(80,2),current_sheet.cell(81,2),current_sheet.cell(82,2),current_sheet.cell(83,2),current_sheet.cell(84,2),current_sheet.cell(85,2),current_sheet.cell(86,2),current_sheet.cell(87,2),current_sheet.cell(88,2),current_sheet.cell(89,2),current_sheet.cell(90,2),current_sheet.cell(93,1),current_sheet.cell(96,1),current_sheet.cell(99,1)]

                    col = 0 # To start from the first column
                    for items in answers_from_sheet:
                        #if items == sheet2_Data[18]:
                        worksheet1.write(row,col,items.value)
                        col += 1
                    row += 1 # To print the next workbook values in the next row

get_ans() # Call function to start execution
#abc = input("=========================PRESS ANY KEY =================")
workbook.close()
