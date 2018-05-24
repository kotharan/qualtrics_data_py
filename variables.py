
#using this, we can just edit this file instead of haphazardly editing the main
#the fist number in the list is where it is, the second is where its going, a '-1' indicates that I don't know where its going.
#all variables consider column A = 0
checkvariable='check variable is good'
csvheader=['ESDname','number of schools in district','School Name','ODS Provider','Other ODS option','Estimated Cost From School','Grade 5 Students','Grade 6 Students','days and nights','total price per student, calculated from school data','Is the money enough','If less than threshold','ODS Provider Fees','Transportation','Personell Stipends',r'Program Costs (w/o admin fees)','total cost - PRELIMINARY']#'Extra Details from Submitter'
#              0           1                            2                    3                   4                             5              6                  7                               8                                             9                      10                   11                     12                  13                            14
#mark nils as '<NULL>'
#ESD info
ESDname=[17,0] #in
ESD_mailing_st=[18,0]
ESD_mailing_City=[19,-1]
ESD_mailing_zip=[20,-1]
ESD_NUM_OF_SCHOOLS=[23,-1]
#number of schools applying q6 #in


#FIRST school info
SCHname=[38,1] #in
hasSCHparticipatedbefore=[39,-1] #in
submitter_contacted_all_schools=[22,-1]

#FIRST outdoor school info
ODS_provider_name=[43,2] #in
ODSotheroption=[44,3]
Facility_name=[45,3] #in
Facility_closest_town=[46,-1] #in


#FIRST program info
grade5_attending=[40,5] #in
grade6_attending=[41,6] #in
program_days_nights=[47,7]  #in

#FIRST program cost
per_student_cost=[0,5 ]#we should add more columns for extra student costs and stuff
per_student_SCH_EST_cost=[50,4] #school estimated cost
SCH_is_money_enough=[48,-1] #surprisingly, our data shows it rarely is #in
SCH_if_threshhold_enough_what_is_needed=[49,-1]
more_money_needed_reason=[50,-1] #in
more_money_needed_reason_othertext=[51,-1] #in
extra_funds_needed_raising_method=[52,-1]
extra_funds_needed_raising_method_othertext=[53,-1]

#FIRST funds breakdown
ODS_provider_fees=[55,-1] #in
transport=[56,-1]#in
personnel_stipends=[57,-1]#in
programcosts_no_admin_costs=[58,-1]#in
details_text_box=[60,-1]#in

#number 59, submitter comments Q30 #in


# submitter info
submitter_name_last=[9,-1]
submitter_name_first=[10,-1]
submitter_email=[11,-1]
submitter_zip=[20,-1]
submitter_agrees_whowhatwhy=[21,-1]
