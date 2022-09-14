#Mukul 2001CE35

#Importing pandas.
import pandas as pAnDa_jI

#Creating a funtion to categorise given data to their respective octants.
def oCtAnT(i,j,k):
    #Various conditions to allocate octants.
    if i>0 and j>0 and k>0:
        return 1
    elif i>0 and j>0 and k<0:
        return -1
    elif i<0 and j>0 and k>0:
        return 2
    elif i<0 and j>0 and k<0:
        return -2
    elif i<0 and j<0 and k>0:
        return 3
    elif i<0 and j<0 and k<0:
        return -3
    elif i>0 and j<0 and k>0:
        return 4
    else :
        return -4
#Here it will read input file.
DaTaFrAmE=pAnDa_jI.read_csv("octant_input.csv")

#Pre-Processing the data.
DaTaFrAmE.at[0,'U_AVG']=DaTaFrAmE['U'].mean()
DaTaFrAmE.at[0,'V_AVG']=DaTaFrAmE['V'].mean()
DaTaFrAmE.at[0,'W_AVG']=DaTaFrAmE['W'].mean()

DaTaFrAmE['U-U_AVG']=DaTaFrAmE['U']-DaTaFrAmE.at[0,'U_AVG']
DaTaFrAmE['V-V_AVG']=DaTaFrAmE['V']-DaTaFrAmE.at[0,'V_AVG']
DaTaFrAmE['W-W_AVG']=DaTaFrAmE['W']-DaTaFrAmE.at[0,'W_AVG']

#Using "oCtAnT" func. for categorizing the data using .apply function.
DaTaFrAmE['octant']=DaTaFrAmE.apply(lambda i: oCtAnT(i['U-U_AVG'], i['V-V_AVG'], i['W-W_AVG']),axis=1)

#Adding/Leaving empty column.
DaTaFrAmE[' ']=''
DaTaFrAmE.at[1,' ']='user input'

#Here we will count overall values using value-counts function.
DaTaFrAmE.at[0,'octant ID']='overall count'
DaTaFrAmE.at[0,'1']=DaTaFrAmE['octant'].value_counts()[1]
DaTaFrAmE.at[0,'-1']=DaTaFrAmE['octant'].value_counts()[-1]
DaTaFrAmE.at[0,'2']=DaTaFrAmE['octant'].value_counts()[2]
DaTaFrAmE.at[0,'-2']=DaTaFrAmE['octant'].value_counts()[-2]
DaTaFrAmE.at[0,'3']=DaTaFrAmE['octant'].value_counts()[3]
DaTaFrAmE.at[0,'-3']=DaTaFrAmE['octant'].value_counts()[-3]
DaTaFrAmE.at[0,'4']=DaTaFrAmE['octant'].value_counts()[4]
DaTaFrAmE.at[0,'-4']=DaTaFrAmE['octant'].value_counts()[-4]

#Taking input from user.
mod=int(input("Please enter the value: "))
DaTaFrAmE.at[1,'octant ID']=mod

total_len=len(DaTaFrAmE['octant'])
temp1=0
