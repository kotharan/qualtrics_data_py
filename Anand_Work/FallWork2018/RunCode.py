'''
 Done By >> Anand Shantilal Kothari 
 For     >> Outdoor Schooling Department at Oregon State University.
'''

import glob                                                  # To search for extension
import xlrd                                                  # To work with spreadsheet/excel data
import os, fnmatch                                           # To work with file location
import xlsxwriter
from itertools import chain

#Pull out working dir path
dir_path = os.path.dirname(os.path.realpath(__file__))

inputbook = xlrd.open_workbook(dir_path+"\SaveAs__.XLSX__Report\\Provider_survey_data.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE OUTPUT======================#

output_Workbook = xlsxwriter.Workbook(dir_path+'\OUTPUT.xlsx')                             #Creating new output_Workbook to write into
outWorkSheet = output_Workbook.add_worksheet('Provider Suvery Data')                       #Creating new worksheet to write into

# Used only the bold ones to see the facilitynames and their other data other than the first facility data
italic = output_Workbook.add_format({'italic': True}) #To do cell text formatting
bold = output_Workbook.add_format({'bold': True})

# For headings

Section_4Head = 49                                                                         # For the headings of the next section after a gap of data 
Section_5Head = 70                                                                         # For the headings of the next section after a gap of data
Section_6Head = 84                                                                         # For the headings of the next section after a gap of data
Section_7Head = 94                                                                         # For the headings of the next section after a gap of data
Section_8Head = 99
Section_9Head = 109
Section_10Head = 126


for headcol in chain(range(0,49),range(139,160),range(349,363),range(1735,1745),range(1835,1840),range(1885,1895),range(1985,2002),range(2155,2158)):    # Since we are taking heads after a gap of data I have specified different ranges 
    head_obj = worksheet.cell(1, headcol)                                                  # Get cell object by row, col
    if headcol <=49:                                                                       # For 1st section of data the headings should be written
        outWorkSheet.write(0,headcol,head_obj.value)
    elif headcol>49 and headcol<=159:                                                      
                                                                                           # For next section heading should be written continuously after the first set of heading avoding the loop headings
        outWorkSheet.write(0,Section_4Head,head_obj.value)
        Section_4Head+=1
    elif headcol>160 and headcol<=362:
        outWorkSheet.write(0,Section_5Head,head_obj.value)
        Section_5Head+=1
    elif headcol>362 and headcol<1745:
        outWorkSheet.write(0,Section_6Head,head_obj.value)
        Section_6Head+=1
    elif headcol>1745 and headcol<1840:
        outWorkSheet.write(0,Section_7Head,head_obj.value)
        Section_7Head+=1
    elif headcol>1840 and headcol<1895:
        outWorkSheet.write(0,Section_8Head,head_obj.value)
        Section_8Head+=1
    elif headcol>1895 and headcol<2002:
        outWorkSheet.write(0,Section_9Head,head_obj.value)
        Section_9Head+=1
    elif headcol>2002 and headcol<2158:
        outWorkSheet.write(0,Section_10Head,head_obj.value)
        Section_10Head+=1
    

"""
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Functions =>
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

SectionNUM(Parameters)  ::>     This starts reading from the 3rd row of the input xlrd data and keeps on looping till the last column/row.
                                Then it checks for a cell which speciefies the number_of_facilities and loops for each facility's
                                data and prints it one below the other. After it is done looping through all the facilities it also creates an
                                empty row so the next data is easily identifyable. [It might be hard to understand by this descriptiton but it makes more sense when you compare the input and output file. Input file is stored in "SaveAs__.XLSX__Report" Directory]


---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Parameters =>
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
- Read_Start_from_row   ::>     Point from where it starts reading rows from the input file.
- till_row              ::>     Reads till the end of the document
- Read_Start_from_col   ::>     Point from where it starts reading columns from the input file.
- till_col              ::>     Reads till the specified column
- Write_Start_from_row  ::>     Specifies the cell to start writing in the OUTPUT file
- Write_Start_from_col  ::>     Specifies the cell to start writing in the OUTPUT file

"""


def Section3b(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=39                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 39 + (10 * num_of_facilities)

            # This loops as per the number of questions for each facility i.e the variable "num_of_facilities"
            for col in range(49,Facility_Row_End):                                 
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 49:
                    Write_Start_from_col = 39
                    Write_Start_from_row += 1

        Write_Start_from_col=0                                                         # Start from the col 0 after you print data of a ?
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section4(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=49                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 139 + (21 * num_of_facilities)                          # 139 is the start of read col 

                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(160,Facility_Row_End):                               
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 70:                                         # 70 is end col of 1st loop
                    Write_Start_from_col = 49
                    Write_Start_from_row += 1

        Write_Start_from_col=49                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section5(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=70                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 349 + (14 * num_of_facilities)                          # This specifies where to end the inner facilities loop as per the "num_of_facilities" 

                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(363,Facility_Row_End):                               
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col                
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 84:                                         # End col of 1st loop
                    Write_Start_from_col = 70
                    Write_Start_from_row += 1

        Write_Start_from_col=70                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section6(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=84                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1735 + (10 * num_of_facilities)                         # This specifies where to end the inner facilities loop as per the "num_of_facilities" 
            
                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(1745,Facility_Row_End):                                
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 94:                                         # End of 1st loop for that facility
                    Write_Start_from_col = 84
                    Write_Start_from_row += 1

        Write_Start_from_col=84                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section7(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=94                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1835 + (5 * num_of_facilities)                          # This specifies where to end the inner facilities loop as per the "num_of_facilities" 
            
                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(1840,Facility_Row_End):                                
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 99:                                         # End of 1st loop for that facility
                    Write_Start_from_col = 94
                    Write_Start_from_row += 1

        Write_Start_from_col=94                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section8(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=99                                                    # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1885 + (10 * num_of_facilities)                         # This specifies where to end the inner facilities loop as per the "num_of_facilities" 
            
                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(1895,Facility_Row_End):                                
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 109:                                        # End of 1st loop for that facility
                    Write_Start_from_col = 99
                    Write_Start_from_row += 1

        Write_Start_from_col=99                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section9(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=109                                                   # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1985 + (17 * num_of_facilities)                         # This specifies where to end the inner facilities loop as per the "num_of_facilities" 
            
                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(2002,Facility_Row_End):                                
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 126:                                        # End of 1st loop for that facility
                    Write_Start_from_col = 109
                    Write_Start_from_row += 1

        Write_Start_from_col=109                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                         # Go to the next row as well

def Section10(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                                     # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                  # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                      
            # Loop for this section in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                     # Stores the Number of Facilities 
            Write_Start_from_col=126                                                   # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 2155 + (3 * num_of_facilities)                          # This specifies where to end the inner facilities loop as per the "num_of_facilities" 
            
                                                                                       # This loops as per the number of questions for each facility i.e the variable "num_of_facilities" and first number is the start of second loop
            for col in range(2158,Facility_Row_End):                                
                cell_data = worksheet.cell(row, col).value                             # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 129:                                        # End of 1st loop for that facility
                    Write_Start_from_col = 126
                    Write_Start_from_row += 1

        Write_Start_from_col=126                                                        # Start from this col after the facility data is printed
        Write_Start_from_row+=1                                                         # Go to the next row as well



Section3b(3,worksheet.nrows,0,49,1,0)                                                   # Calling then function Section3b
Section4(3,worksheet.nrows,139,160,1,49)                                                # Calling then function Section4 and so on
Section5(3,worksheet.nrows,349,363,1,70)
Section6(3,worksheet.nrows,1735,1745,1,84)
Section7(3,worksheet.nrows,1835,1840,1,94)
Section8(3,worksheet.nrows,1885,1895,1,99)
Section9(3,worksheet.nrows,1985,2002,1,109)
Section10(3,worksheet.nrows,2155,2158,1,126)

output_Workbook.close()                                                                # Close the workbook when done
