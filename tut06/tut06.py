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
    
    #Creating a dataframe to store to consolidated output.
    FIN_DaTaFrAmE = pAnDa_jI.DataFrame()
    FIN_DaTaFrAmE['Roll'] = ''
    FIN_DaTaFrAmE['Name'] = ''
    for i in range(0, len(VAL_DATES)):
        FIN_DaTaFrAmE[VAL_DATES[i]] = ''
    FIN_DaTaFrAmE['Actual Lecture Taken'] = ''
    FIN_DaTaFrAmE['Total Real'] = ''
    FIN_DaTaFrAmE['% Attendance'] = ''

    #Giving structure according to our valid dates stored in the list.
    for i in DaTaFrAmE2.index:
        FIN_DaTaFrAmE.at[i+1, 'Roll'] = DaTaFrAmE2['Roll No'][i]
        FIN_DaTaFrAmE.at[i+1, 'Name'] = DaTaFrAmE2['Name'][i]
        FIN_DaTaFrAmE.at[i+1, 'Actual Lecture Taken'] = len(VAL_DATES)
        FIN_DaTaFrAmE.at[i+1,
                         'Total Real'] = FIN_DaTaFrAmE.at[i+1, '% Attendance'] = 0

    try:
        #Main Loop (iterating for each roll number one by one)
        for i in DaTaFrAmE2.index:
            #A variable dataframe which will store data for every roll number in each iteration.
            DaTaFrAmE3 = pAnDa_jI.DataFrame()
            DaTaFrAmE3['Date'] = ''
            DaTaFrAmE3['Roll'] = ''
            DaTaFrAmE3['Name'] = ''
            DaTaFrAmE3['Total Attendance Count'] = 0
            DaTaFrAmE3['Real'] = ''
            DaTaFrAmE3['Duplicate'] = ''
            DaTaFrAmE3['Invalid'] = ''
            DaTaFrAmE3['Absent'] = ''

            DaTaFrAmE3.at[0, 'Roll'] = DaTaFrAmE2['Roll No'][i]
            DaTaFrAmE3.at[0, 'Name'] = DaTaFrAmE2['Name'][i]

            #Writing valid dates in the dataframe from previously formed list.
            for j in range(1, len(VAL_DATES)+1):
                DaTaFrAmE3.at[j, 'Date'] = VAL_DATES[j-1]
                DaTaFrAmE3.at[j, 'Total Attendance Count'] = DaTaFrAmE3.at[j, 'Real'] = DaTaFrAmE3.at[j,
                                                                                                      'Duplicate'] = DaTaFrAmE3.at[j, 'Invalid'] = DaTaFrAmE3.at[j, 'Absent'] = 0

            #Temporary variable
            temp = DaTaFrAmE2['Roll No'][i]

            # Counting attendance according to the given criteria and storing in respective columns.
            for j in DaTaFrAmE1.index:
                if (temp == DaTaFrAmE1['Roll'][j]):
                    if (DaTaFrAmE1['Date'][j] != -1):
                        ind = VAL_DATES.index(DaTaFrAmE1['Date'][j])
                        DaTaFrAmE3.at[ind+1, 'Total Attendance Count'] += 1
                        if (DaTaFrAmE1['Time'][j] >= '14:00' and DaTaFrAmE1['Time'][j] <= '15:00'):
                            if (DaTaFrAmE3['Real'][ind+1] == 0):
                                DaTaFrAmE3.at[ind+1, 'Real'] += 1
                            else:
                                DaTaFrAmE3.at[ind+1, 'Duplicate'] += 1
                        else:
                            DaTaFrAmE3.at[ind+1, 'Invalid'] += 1

            # Marking absent from the above evaluated attendance.
            for j in range(1, len(VAL_DATES)+1):
                if (DaTaFrAmE3['Real'][j] == 0):
                    DaTaFrAmE3.at[j, 'Absent'] = 1
                    FIN_DaTaFrAmE.at[i+1, VAL_DATES[j-1]] = 'A'
                else:
                    FIN_DaTaFrAmE.at[i+1, VAL_DATES[j-1]] = 'P'
                    FIN_DaTaFrAmE.at[i+1, 'Total Real'] += 1
                FIN_DaTaFrAmE.at[i+1, '% Attendance'] = round(
                    FIN_DaTaFrAmE.at[i+1, 'Total Real']*100/FIN_DaTaFrAmE.at[i+1, 'Actual Lecture Taken'], 2)

            # Saving the excel file for each roll number.
            DaTaFrAmE3.to_excel('output/'+temp+'.xlsx', index=False)
    except:
        print("Index overflow, check the range again.")
