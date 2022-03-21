# -*- coding: utf-8 -*-
"""
Created on Mon Dec 27 13:46:05 2021

@author: TessJanney
"""


import pandas as pd
from numpy import nan as Nan
import PySimpleGUI as sg



#Importing excel file 


#df = pd.read_excel(r'C:\Users\TessJanney\CODE\Timecard Automation\coins_weekly_report (40)_mkj.xlsx')

df = pd.read_excel('coins_weekly_report (40)_mkj.xlsx')


#This printed out all the rows from column 2 of the input excel sheet 
#projnames = df.iloc[:,2]

#creating the data frame that contains all the columns I need from the input excel sheet. 
df[["first_name","last_name","tce_date","project_number","project_name","tce_rhrs","notes","position_name" ]]


#This is the output file name and where the output file is written. It is as the top of the code because if it is at the 
#bottom, the file continues to  be overwritten with each loop and only the last loop is printed. 
file_name =  'Timecards2.xlsx'
writefile = pd.ExcelWriter(file_name,engine='xlsxwriter')

#This loop is looping through all the project names from the original data frame and identifying the names of the poeple 
#who worked on those projects and the hours they spent on them. The total hours for each person on each project are summed. 
#The final timecards for each project are then sorted into different sheets depending on the project name.  
df_all = pd.DataFrame()
for project_name in df.iloc[:,2].unique():
    #print(project_name)
    df_temp1 = df[df['project_name']==project_name]
    

#identifying which headers I want to keep for the output file
    df_final = df_temp1[["first_name","last_name","tce_date","project_number","project_name","tce_rhrs","notes" ,"position_name"]]

#renaming those column headers to more formal formatting
    df_final.rename(columns = {'first_name':'First Name','last_name':'Last Name','tce_date':'Date','project_number':'Project Number',
                           'project_name':'Project Name','tce_rhrs':'Hours','notes':'Description','position_name':'Position Name'}, inplace = True)
    df_hours = pd.DataFrame()

    for first_name in df_final.iloc[:,0].unique():
        #print(first_name)
        df_temp = df_final[df_final['First Name']==first_name]
        hours_total = df_temp['Hours'].sum()
        r1 = pd.Series([Nan,Nan,Nan,Nan,Nan,Nan,Nan], index = ["First Name","Last Name","Date","Project Number","Project Name","Hours","Description"])
    
        df_hours = df_hours.append(df_temp)
        
   
        total = "Total"
        df_hours = df_hours.append({'Hours': hours_total,'Project Name': total}, ignore_index=True)
        df_hours = df_hours.append(r1, ignore_index=True)
        df_hours = df_hours.append(r1, ignore_index=True)
        
    #Creating parameters for the output file.    
        project_number = tuple(df_hours['Project Number'])
        #print(project_number)
        hash(project_number)
    sheet_name = project_number[0]
    print(sheet_name)
   #Excel can only write certain characters, so this is replacing the non useable characters with useable ones.  
    sheet_name = str(sheet_name) 
    if len(sheet_name)>30:
        sheet_name = sheet_name[:30]
    if '/' in sheet_name:
        sheet_name = sheet_name.replace('/', '-')
    if ':' in sheet_name:
        sheet_name= sheet_name.replace(':', '-')
    
    #print(sheet_name)
    
    df_hours.to_excel(writefile, sheet_name=sheet_name, index =
                  False, startrow = 1)

    workbook = writefile.book 
    worksheet = writefile.sheets[sheet_name]
    worksheet.write(0, 6,'Employee Project Time Cards', workbook.add_format
              ({'bold': True, 'color':'#000000', 'size' : 14}))
    worksheet.insert_image('A1','Alloy Horizontal ORG.jpg',{'x_scale':0.08,'y_scale':0.08})
    worksheet.set_column('A:F', 20)
    worksheet.set_row(0,60)    

    #r1 = pd.Series([Nan,Nan,Nan,Nan,Nan,Nan,Nan], index = ["First Name","Last Name","Date","Project Number","Project Name","Hours","Description"])        
    #df_all = df_all.append(df_hours)
    #df_all = df_all.append(r1, ignore_index=True)
    #df_all = df_all.append(r1, ignore_index=True)




writefile.save()


#print()


    

