import csv
from time import sleep
from variables import *  #I offloaded all our column nmaes into another script file

print(checkvariable)

with open('outdoor.csv',newline='') as f:
    reader = csv.reader(f)
    row1 = next(reader) # gets the first line
    row2 = next(reader) # the second
     # now do something here
    # if first row is the header, then you can do one more next() to get the next row:
    # row2 = next(f)
for k in range(len(row2)):
    print(row2[k]) #fyi python uses brackets for list indices
    sleep(1)
