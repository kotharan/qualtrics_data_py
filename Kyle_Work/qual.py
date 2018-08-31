from time import sleep, time
t1=time()
import csv
from variables import *  #I offloaded all our column nmaes into another script file
from create_single_line import * #theres some extra functions in this file as well

#unholy global variables
slptm=3
skipnum=4
list_of_schools=[]


#open files
with open('outdoor2.csv',newline='') as f, open('output.csv','w',newline='') as b:
    reader = csv.reader(f)
    write = csv.writer(b)
    write.writerow(csvheader)
    #write.writerow([''])  #the fancy space for readability
    for k in list(range(skipnum)): #this section skips past all the labels and gets right to the data
        row=next(reader)
    while True:  #goes row by row reading each individual school's data
        num=get_num_of_schools_row(row) #get num of schools in row
        #print('total schools: '+str(num))
        for k in list(range(num)):
            rowscraped=scrape_data_into_list(row,k+1) #get each schools info
            list_of_schools.append(rowscraped[2])
            #print(str(rowscraped))
            write.writerow(rowscraped) #write that info to a new line in the spreadsheet
        #write.writerow([''])  per school line break
        try: #this will stop the program when its out of data
            row=next(reader)
        except StopIteration:
            break
t2=time() #this tells us how long it took (never more than 200ms)
ttl=round((t2-t1)*1000,3)

#print('\n\n',str(list_of_schools),'\n\n')
print('finished in '+str(ttl)+ ' milliseconds! Exiting in '+str(slptm)+' seconds...')
sleep(slptm)
