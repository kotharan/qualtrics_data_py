import csv
from time import sleep
from variables import *  #I offloaded all our column nmaes into another script file
from create_single_line import *

print(checkvariable)

with open('outdoor.csv',newline='') as f:
    reader = csv.reader(f)
    skipnum=int(input('choose row number: '))
    for k in list(range(skipnum)):
        row=next(reader)


    num=get_num_of_schools_row(row)
    print('total schools: '+str(num))
    for k in list(range(num)):
        rowscraped=scrape_data_into_list(row,k+1)
        print(str(rowscraped))

    #print(str(row))
