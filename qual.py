import csv
import pandas as pd

with open('outdoor.csv','rb') as file:
    # reader = csv.reader(file)
    # lines = [line for line in file]
    # print(lines[25])


    df = pd.read_csv(file)
    saved_column = df.EndDate + "   " + df.StartDate #you can also use df['column_name']

    print(saved_column)
    a = list(df.Progress[2:].astype(int))
    print(sum(a))
