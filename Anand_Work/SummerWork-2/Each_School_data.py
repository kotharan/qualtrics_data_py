import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter

#Pull out working dir path
dir_path = os.path.dirname(os.path.realpath(__file__))

inputbook = xlrd.open_workbook(dir_path+"\outdoor2.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE======================#

output_Workbook = xlsxwriter.Workbook(dir_path+'\OUTPUT.xlsx') #Creating new output_Workbook to write into
outWorkSheet = output_Workbook.add_worksheet('All Schools Info') #Creating new worksheet to write into

# Used only the bold ones to see the schoolnames and their other data other than the first school data
italic = output_Workbook.add_format({'italic': True}) #To do cell text formatting
bold = output_Workbook.add_format({'bold': True})


# For headings
for headcol in range(0,62):
    head_obj = worksheet.cell(1, headcol)  # Get cell object by row, col
    outWorkSheet.write(0,headcol,head_obj.value)


# For printing data
"""
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Functions =>
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

DataLoop(Parameters)              ::>     This starts reading from the 3rd row of the input xlrd data and keeps on looping till the 62th column.
                                Then it checks for a cell which speciefies the number of school that district has and loops to each school's
                                data and prints in under the main district row. After it is done looping through all the schools it also creates an
                                empty row so the other district data is easily identifyable. It also applies bold cell format to each school data after
                                the first main district row data. [It might be hard to understand by this descriptiton but it makes more sense when your compare
                                the input and output file]


---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Parameters =>
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
- Read_Start_from_row   ::>     Point from where it starts reading rows from the input file.
- till_row              ::>     Reads till the end of the document
- Read_Start_from_col   ::>     Point from where it starts reading columns from the input file.
- till_col              ::>     Reads till the specified column
- Write_Start_from_row  ::>     Point to start writing in the OUTPUT xl file
- Write_Start_from_col  ::>     Point to start writing in the OUTPUT xl file

"""

def DataLoop(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)
            Write_Start_from_col +=1

        if( str(worksheet.cell(row, 23).value) > str(1.0)):                          # This checks if there are more than one school and loops to each school data if there is
            num_of_schools = int(worksheet.cell(row, 23).value)
            Write_Start_from_col=38                                                  # If there is more than one school print the next school data below the first school data
            Write_Start_from_row+=1
            schoolData_Row_End = 29*(num_of_schools+1)

            for col in range(62,schoolData_Row_End):                                 # This is 29 times num_of_schools of schools because there are 29 questions for each school
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data,bold)
                Write_Start_from_col +=1

                if Write_Start_from_col == 62:
                    Write_Start_from_col = 38
                    Write_Start_from_row += 1

        Write_Start_from_col=0                                                         # Start from the col 0 after you print data of a district
        Write_Start_from_row+=1                                                        # Go to the next row as well



DataLoop(3,worksheet.nrows,0,62,1,0)                                                   # Calling then function DataLoop

output_Workbook.close()                                                                # Close the workbook when done
