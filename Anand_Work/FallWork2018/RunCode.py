import glob # To search for extension
import xlrd # To work with spreadsheet/excel data
import os, fnmatch # To work with file location
import xlsxwriter
from itertools import chain

#Pull out working dir path
dir_path = os.path.dirname(os.path.realpath(__file__))

inputbook = xlrd.open_workbook(dir_path+"\\Provider_survey_data.xlsx") # Book is the name of the workbook of the raw data and it stores the running workbook whose value would be taken as isput from the raw files
worksheet = inputbook.sheet_by_index(0)

#===========CREATING A WORKBOOK TO STORE======================#

output_Workbook = xlsxwriter.Workbook(dir_path+'\OUTPUT.xlsx') #Creating new output_Workbook to write into
outWorkSheet = output_Workbook.add_worksheet('Provider Suvery Data') #Creating new worksheet to write into

# Used only the bold ones to see the facilitynames and their other data other than the first facility data
italic = output_Workbook.add_format({'italic': True}) #To do cell text formatting
bold = output_Workbook.add_format({'bold': True})

words2words1 = {'One': 1, 'Two': 2, 'Three' : 3, 'Four': 4, 'Five': 5, \
            'Six': 6, 'Seven': 7, 'Eight': 8 , 'Nine': 9,'Ten': 10 }

# For headings

Section_4Head = 49                                      # For the headings of the next section after a gap of data 
Section_5Head = 70                                      # For the headings of the next section after a gap of data
Section_6Head = 84                                      # For the headings of the next section after a gap of data
Section_7Head = 94                                      # For the headings of the next section after a gap of data
Section_8Head = 99
Section_9Head = 109
Section_10Head = 126


for headcol in chain(range(0,49),range(139,160),range(349,363),range(1735,1745),range(1835,1840),range(1885,1895),range(1985,2002),range(2155,2158)):    # Since we are taking heads after a gap of data I have specified different ranges 
    head_obj = worksheet.cell(1, headcol)            # Get cell object by row, col
    if headcol <=49:                                 # For 1st section of data the headings should be written
        outWorkSheet.write(0,headcol,head_obj.value)
    elif headcol>49 and headcol<=159:                # For next section heading should be written continuously after the first set of heading avoding the loop headings
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
    


# For printing data
"""
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Functions =>
---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Section3b(Parameters)              ::>     This starts reading from the 3rd row of the input xlrd data and keeps on looping till the 62th column.
                                Then it checks for a cell which speciefies the number of facilities that ? has and loops to each facility's
                                data and prints in under the main ? row. After it is done looping through all the facilitys it also creates an
                                empty row so the other ? data is easily identifyable. It also applies bold cell format to each facility data after
                                the first main ? row data. [It might be hard to understand by this descriptiton but it makes more sense when your compare
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


def Section3b(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=39                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 39 + (10 * num_of_facilities)

            for col in range(49,Facility_Row_End):                                 # This is 29 times num_of_facilities of facility because there are 29 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 49:
                    Write_Start_from_col = 39
                    Write_Start_from_row += 1

        Write_Start_from_col=0                                                         # Start from the col 0 after you print data of a ?
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section4(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=49                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 139 + (21 * num_of_facilities)                       # 138 is 1 minus of starting read col 

            for col in range(160,Facility_Row_End):                                #160 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 70:                                      # 70 is end col of 1st loop
                    Write_Start_from_col = 49
                    Write_Start_from_row += 1

        Write_Start_from_col=49                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section5(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=70                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 349 + (14 * num_of_facilities)                       # 139 of starting read col 

            for col in range(363,Facility_Row_End):                                #160 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col                
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 84:                                      # 70 is end col of 1st loop
                    Write_Start_from_col = 70
                    Write_Start_from_row += 1

        Write_Start_from_col=70                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section6(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=84                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1735 + (10 * num_of_facilities)                       # 139 of starting read col 

            for col in range(1745,Facility_Row_End):                                #1745 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 94:                                      # 94 is end col of 1st loop
                    Write_Start_from_col = 84
                    Write_Start_from_row += 1

        Write_Start_from_col=84                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section7(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=94                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1835 + (5 * num_of_facilities)                       # 139 of starting read col 

            for col in range(1840,Facility_Row_End):                                #1745 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 99:                                      # 94 is end col of 1st loop
                    Write_Start_from_col = 94
                    Write_Start_from_row += 1

        Write_Start_from_col=94                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section8(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=99                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1885 + (10 * num_of_facilities)                       # 139 of starting read col 

            for col in range(1895,Facility_Row_End):                                #1745 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 109:                                      # 94 is end col of 1st loop
                    Write_Start_from_col = 99
                    Write_Start_from_row += 1

        Write_Start_from_col=99                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section9(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=109                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 1985 + (17 * num_of_facilities)                       # 139 of starting read col 

            for col in range(2002,Facility_Row_End):                                #1745 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 126:                                      # 94 is end col of 1st loop
                    Write_Start_from_col = 109
                    Write_Start_from_row += 1

        Write_Start_from_col=109                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well

def Section10(Read_Start_from_row, till_row , Read_Start_from_col , till_col , Write_Start_from_row , Write_Start_from_col ):
            # Read of row start from the data file                             # Writing in the OUTPUT file

    for row in range(Read_Start_from_row , till_row):
        for col in range(Read_Start_from_col, till_col):
            cell_obj = worksheet.cell(row, col).value                                # Get cell object by row, col
            outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_obj)     # Write the cell_obj in OUTPUT file
            Write_Start_from_col +=1
        
        if(worksheet.cell(row, 19).value == ""):
            pass
        elif (worksheet.cell(row, 19).value > 1):                                    # Loop for Section 3B in qualtrics. This checks if there are more than one facility and loops to each facility data if there is
            num_of_facilities = int(worksheet.cell(row, 19).value)                   # Stores the Number of Facilities 
            Write_Start_from_col=126                                                  # If there is more than one facility print the next facility data below the first facility data
            Write_Start_from_row+=1
            Facility_Row_End = 2155 + (3 * num_of_facilities)                       # 139 of starting read col 

            for col in range(2158,Facility_Row_End):                                #1745 is second loop start data # This is 21 times num_of_facilities of facility because there are 21 questions for each facility
                cell_data = worksheet.cell(row, col).value                           # Get cell object by row, col
                outWorkSheet.write(Write_Start_from_row,Write_Start_from_col,cell_data)
                Write_Start_from_col +=1

                if Write_Start_from_col == 129:                                      # 94 is end col of 1st loop
                    Write_Start_from_col = 126
                    Write_Start_from_row += 1

        Write_Start_from_col=126                                                        # Start from the col 49 after you print data of a facility
        Write_Start_from_row+=1                                                        # Go to the next row as well



Section3b(3,worksheet.nrows,0,49,1,0)                                                   # Calling then function Section3b
Section4(3,worksheet.nrows,139,160,1,49)                                               # Calling then function Section4
Section5(3,worksheet.nrows,349,363,1,70)
Section6(3,worksheet.nrows,1735,1745,1,84)
Section7(3,worksheet.nrows,1835,1840,1,94)
Section8(3,worksheet.nrows,1885,1895,1,99)
Section9(3,worksheet.nrows,1985,2002,1,109)
Section10(3,worksheet.nrows,2155,2158,1,126)

output_Workbook.close()                                                                # Close the workbook when done
