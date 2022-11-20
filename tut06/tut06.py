#Mukul 2001CE35
#Importing pandas.

import pandas as pAnDa_jI

from datetime import datetime
start_time = datetime.now()

#In the below code we will be reading and writing in multiple .xlsx files.

def AtTeNdAnCe_ReP():
    try:
        #Reading the input .csv file and storing it in a variable.
        DaTaFrAmE1=pAnDa_jI.read_csv('input_attendance.csv')
        DaTaFrAmE2=pAnDa_jI.read_csv('input_registered_students.csv')
    except:
        print("There was an error reading the file.")
    #Adding empty columns for processesing the given data.
    DaTaFrAmE1['Roll']=''
    DaTaFrAmE1['Time']=''
    DaTaFrAmE1['Date']=''
    DaTaFrAmE1['Day']=''
    DaTaFrAmE1['Month']=''
    DaTaFrAmE1['Year']=''
    
    #Concatenating the given data into desired format and string in respective columns.
    DaTaFrAmE1.loc[DaTaFrAmE1.Roll=='','Roll']=DaTaFrAmE1.Attendance.str.split().str.get(0)
    DaTaFrAmE1.loc[DaTaFrAmE1.Time=='','Time']=DaTaFrAmE1.Timestamp.str.split().str.get(1)
    DaTaFrAmE1.loc[DaTaFrAmE1.Date=='','Date']=DaTaFrAmE1.Timestamp.str.split().str.get(0)
    DaTaFrAmE1.loc[DaTaFrAmE1.Day=='','Day']=DaTaFrAmE1.Date.str.split('-').str.get(0)
    DaTaFrAmE1.loc[DaTaFrAmE1.Month=='','Month']=DaTaFrAmE1.Date.str.split('-').str.get(1)
    DaTaFrAmE1.loc[DaTaFrAmE1.Year=='','Year']=DaTaFrAmE1.Date.str.split('-').str.get(2)

    #An empty list to store the valid days (Modays and Thurdays)
    VAL_DATES=[]
    SaKuRa=''
    #Indexing over the dataframe for getting valid dates, using 'datetime' library for the same.
    for i in DaTaFrAmE1.index:
        DATE = datetime(int(DaTaFrAmE1['Year'][i]),int(DaTaFrAmE1['Month'][i]),int(DaTaFrAmE1['Day'][i]))
        if (DATE.weekday()==0 or DATE.weekday()==3):
            if (SaKuRa!= DaTaFrAmE1['Date'][i]):
                VAL_DATES.append(DaTaFrAmE1['Date'][i])
                SaKuRa = DaTaFrAmE1['Date'][i]
        else:
            DaTaFrAmE1['Date'][i]=-1

    #Sorting our processed dataframe according to roll numbers for the ease of evaluation.
    DaTaFrAmE1=DaTaFrAmE1.sort_values('Roll')