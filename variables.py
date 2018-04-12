#using this, we can just edit this file instead of haphazardly editing the main
#the fist number in the list is where it is, the second is where its going, a '-1' indicates that I don't know where its going.
#all variables consider column A = 0
checkvariable='check variable is good'

#ESD info
ESDname=[17,0]
ESD_mailing_st=[18,0]
ESD_mailing_City=[19,-1]
ESD_mailing_zip=[20,-1]



#FIRST school info
SCHname=[39,1]
hasSCHparticipatedbefore=[40,-1]


#FIRST outdoor school info
ODS_provider_name=[44,2]
ODSotheroption=[45,-1]
Facility_name=[47,3]
Facility_closest_town=[47,-1]


#FIRST program info
grade5_attending=[41,4]
grade6_attending=[42,5]
program_days&nights=[48,6]

#FIRST program cost
per_student_cost=[0,8 ]#we should add more columns for extra student costs and stuff
per_student_SCH_EST_cost=[50,9] #school estimated cost
SCH_is_money_enough=[49,-1] #surprisingly, our data shows it rarely is
more_money_needed_reason=[51,-1]
more_money_needed_reason_othertext=[52,-1]
extra_funds_needed_raising_method=[53,-1]
extra_funds_needed_raising_method_othertext=[54,-1]

#funds breakdown
ODS_provider_fees=[56,-1]
transport=[57,-1]
personnel_stipends=[58,-1]
programcosts_no_admin_costs=[59,-1]
details_text_box=[61,-1]




# submitter info
submitter_name_last=[9,-1]
submitter_name_first=[10,-1]
submitter_email=[11,-1]
submitter_zip=[20,-1]
submitter_agrees_whowhatwhy=[21,-1]
submitter_contacted_all_schools=[22,-1]
