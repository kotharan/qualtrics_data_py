import glob
import csv
import time


starttime=time.time()
filepath= r'E:\code\qualtrics_data_py\stage 2\dump reports here\*'
outputfile= r'E:\code\qualtrics_data_py\stage 2\results.csv'


fn=glob.glob(filepath)
# stt=[' '*30]
# for k in range(len(fn)):
#     stt[((k)*2)]=fn[k]
#     stt[(k)*2+1]='\n'
print(str(len(fn)))
print(str(fn))

# for fnnumber in len(fn):
#     with open(fn(fnnumber),newline='') as f, open('output.csv','w',newline='') as b:
#         reader = csv.reader(f)
#         row=next(reader)
#         print(row)
