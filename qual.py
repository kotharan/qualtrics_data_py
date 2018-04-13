import csv
from time import sleep
from variables import *  #I offloaded all our column nmaes into another script file
from create_single_line import *

print(checkvariable)

with open('outdoor.csv',newline='') as f:
    reader = csv.reader(f)
    for k in list(range(4)):
        row=next(reader)

    row=scrape_data_into_list(row,SCHname[0])
    print(str(row))
