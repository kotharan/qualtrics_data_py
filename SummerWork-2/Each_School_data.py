import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter


inputbook = xlrd.open_workbook(r"C:\Users\DELL\Documents\GitHub\qualtrics_data_py\SummerWork-2\outdoor2.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE======================#

output_Workbook = xlsxwriter.Workbook('.\SummerWork-2\OUTPUT.xlsx') #Creating new output_Workbook to write into
outWorkSheet = output_Workbook.add_worksheet('All Schools Info') #Creating new worksheet to write into


italic = output_Workbook.add_format({'italic': True}) #To do cell text formatting
bold = output_Workbook.add_format({'bold': True})


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
        if( str(worksheet.cell(row, 23).value) > str(1.0)):
            num_of_schools = int(worksheet.cell(row, 23).value)
            Write_Start_from_col=38
            Write_Start_from_row+=1
            schoolData_Row_End = 29*(num_of_schools+1)
            for col in range(62,schoolData_Row_End):    # This is 29 times num_of_schools of schools because there are 29 questions for each school
                cell_data = worksheet.cell(row, col).value  # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data,bold)
                Write_Start_from_col +=1
                if Write_Start_from_col == 62:
                    Write_Start_from_col = 38
                    Write_Start_from_row += 1

        Write_Start_from_col=0
        Write_Start_from_row+=1



DataLoop(3,worksheet.nrows,0,62,1,0)

output_Workbook.close()

# For now it is reading till column 62 and then skips the further data of other school and goes to the next line
