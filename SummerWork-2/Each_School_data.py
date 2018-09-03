import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter


inputbook = xlrd.open_workbook(r"C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork-2\test.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE======================#

workbook = xlsxwriter.Workbook('.\SummerWork-2\OUTPUT.xlsx') #Creating new workbook to write into
outWorkSheet = workbook.add_worksheet('All Schools Info') #Creating new worksheet to write into
outWorkSheet1 = workbook.add_worksheet('Just School') #Creating new worksheet to write into


cell_format = workbook.add_format({'color':'red'}) #To do cell text formatting
bold = workbook.add_format({'bold': True})


# For heading

for headcol in range(0,62):
    head_obj = worksheet.cell(1, headcol)  # Get cell object by row, col
    outWorkSheet.write(0,headcol,head_obj.value)


# For printing data


def DataLoop(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)
            Write_Start_from_col +=1
        if(worksheet.cell(row, 23).value > 1):
            num_of_schools = int(worksheet.cell(row, 23).value)
            Write_Start_from_col=38
            Write_Start_from_row+=1
            for col in range(62,24*num_of_schools):    # This is 24 times num_of_schools of schools because there are 24 questions for each school
                cell_data = worksheet.cell(row, col).value  # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1
                if Write_Start_from_col == 62:
                    Write_Start_from_col = 38
                    Write_Start_from_row += 1


        #    OutputSchools_Data(row,int(num_of_schools),1,38)
        Write_Start_from_col=0
        Write_Start_from_row+=1

def OutputSchools_Header():
    cols = 38
    for headcol in range(62,86):
        head_obj = worksheet.cell(1, headcol)  # Get cell object by row, col
        outWorkSheet1.write(0,cols,head_obj.value)
        cols += 1

# def OutputSchools_Data(row,num_of_schools,Write_Start_from_row,Write_Start_from_col):
#     At_62 = 62
#     for col in range(At_62,24**num_of_schools):
#         cell_data = worksheet.cell(row, col).value  # Get cell object by row, col
#         outWorkSheet1.write(Write_Start_from_row,Write_Start_from_col,cell_data)
#         Write_Start_from_col +=1
#
#     Write_Start_from_col=0
#     Write_Start_from_row+=1
#

DataLoop(3,20,0,62,1,0)
OutputSchools_Header()
workbook.close()

# For now it is reading till column 62 and then skips the further data of other school and goes to the next line
