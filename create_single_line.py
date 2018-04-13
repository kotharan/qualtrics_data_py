from variables import *

def allocatelist(var,times):  #this just allocates the list
    var=[]
    for k in range(times):
        var.append('_')
    return var

#this tells us the total per student cost, and works for any school in a row
def generate_cost(fullstring,ODS):
    perSTtotal= fullstring[ODS]+fullstring[ODS+1]+fullstring[ODS+2]+fullstring[ODS+3]
    return perSTtotal


#this function scrapes all the data for a given school in a list
def scrape_data_into_list(fullstring,SCHnamelocation):
#this line allows this function to work for any school in a row
    diff=SCHnamelocation-SCHname[0]
    print(diff)
    var=[]
    line=allocatelist(var,6)
    line[ESDname[1]]=fullstring[ESDname[0]] #one ESD per line so this shouldn't change
    line[SCHname[1]]=fullstring[(SCHname[0]+diff)] #school name
    line[ODS_provider_name[1]]=fullstring[ODS_provider_name[0]+diff] #ods name
    if fullstring[ODSotheroption[0]+diff] == '':
        line[ODSotheroption[1]]='__'
    else:
        line[ODSotheroption[1]]=fullstring[ODSotheroption[0]+diff] #ods other option name,
    line[per_student_SCH_EST_cost[1]]=fullstring[per_student_SCH_EST_cost[0]+diff] #school estimated cost per student
    line[per_student_cost[1]]=generate_cost(fullstring,ODS_provider_fees[0]+diff) #cost per student estimated using the cost breakdown
    return line
