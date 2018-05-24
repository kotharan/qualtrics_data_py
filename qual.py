from time import sleep, time
t1=time()
import csv
from variables import *  #I offloaded all our column nmaes into another script file
from create_single_line import * #theres some extra functions in this file as well

#print(checkvariable)
#open files
with open('outdoor2.csv',newline='') as f, open('output.csv','w',newline='') as b:
    reader = csv.reader(f)
    write = csv.writer(b)
    write.writerow(csvheader)
    #write.writerow([''])  #the fancy space for readability

    #this section skips past all the labels and gets right to the data
    skipnum=4
    for k in list(range(skipnum)):
        row=next(reader)
    nxt=True
    while nxt==True:
        num=get_num_of_schools_row(row)
        #print('total schools: '+str(num))
        for k in list(range(num)):
            rowscraped=scrape_data_into_list(row,k+1)
            #print(str(rowscraped))
            write.writerow(rowscraped)
        #write.writerow([''])
        try:
            row=next(reader)
        except StopIteration:
            break
t2=time()
ttl=round((t2-t1)*1000,3)
slptm=3
print('finished in '+str(ttl)+ ' milliseconds! Exiting in '+str(slptm)+' seconds...')
sleep(slptm)
# >>>>>>> a250741f8fb547dd060532dc95850e7e8a5aefa2
