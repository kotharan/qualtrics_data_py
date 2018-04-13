from variables import *

def allocatelist(var,times):  #this just allocates the list
    var=[]
    for k in range(times):
        var.append('_')
    return var

def get_num_of_schools_row(fullstring):#this figures out how many schools are in a row but scanning until it finds a blank where a school name should be.
    try:
        if fullstring[SCHname[0]] != '':
            num=1
        else:
            num=0
            #separation distance between schools is 230 cells
    except IndexError:
        num=0
        run=0
    sep=23
    run=1
    while run != 0:
        try:
            if fullstring[(SCHname[0]+run*sep)] !='':
                run=run+1
                num=run
            else:
                run=0
        except IndexError:
            num=1
            run=0
    return num



def generate_cost(fullstring,ODS): # (UNUSED) this tells us the total per student cost, and works for any school in a row
    perSTtotal= fullstring[ODS]+fullstring[ODS+1]+fullstring[ODS+2]+fullstring[ODS+3]
    return perSTtotal


def scrape_data_into_list(fullstring,schnum): #this function scrapes all the data for a given school in a list
#this line allows this function to work for any school in a row
    diff=(schnum-1)*23
    var=[]
    line=allocatelist(var,11)
    line[ESDname[1]]=fullstring[ESDname[0]] #one ESD per line so this shouldn't change
    line[SCHname[1]]=fullstring[(SCHname[0]+diff)] #school name
    line[ODS_provider_name[1]]=fullstring[ODS_provider_name[0]+diff] #ods name
    if fullstring[ODSotheroption[0]+diff] == '':
        line[ODSotheroption[1]]=''
    else:
        line[ODSotheroption[1]]=fullstring[ODSotheroption[0]+diff] #ods other option name,
    line[per_student_SCH_EST_cost[1]]=fullstring[per_student_SCH_EST_cost[0]+diff] #school estimated cost per student
    for k in list(range(5)):
        line[per_student_cost[1]+k]=fullstring[ODS_provider_fees[0]+diff+k] #cost per student estimated using the cost breakdown

    return line
