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