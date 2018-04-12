import csv
from time import sleep
from variables import *  #I offloaded all our column nmaes into another script file

print(checkvariable)

with open('outdoor.csv',newline='') as f:
    reader = csv.reader(f)
        for k in range(4):
            row=next(reader)
        row=scrape_data_into_list(row,SCHname[0])


     # now do something here
    # if first row is the header, then you can do one more next() to get the next row:
    # row2 = next(f)
for k in range(len(row)):
    print(row[k]) #fyi python uses brackets for list indices
    sleep(1)
