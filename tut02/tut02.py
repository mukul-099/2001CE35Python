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
DaTaFrAmE=pAnDa_jI.read_excel("input_octant_transition_identify.xlsx")

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
DaTaFrAmE['']=''
DaTaFrAmE.at[1,'']='User Input'

#Here we will count overall values using value-counts function.
DaTaFrAmE.at[0,'Octant ID']='Overall Count'
DaTaFrAmE.at[0,'1']=DaTaFrAmE['octant'].value_counts()[1]
DaTaFrAmE.at[0,'-1']=DaTaFrAmE['octant'].value_counts()[-1]
DaTaFrAmE.at[0,'2']=DaTaFrAmE['octant'].value_counts()[2]
DaTaFrAmE.at[0,'-2']=DaTaFrAmE['octant'].value_counts()[-2]
DaTaFrAmE.at[0,'3']=DaTaFrAmE['octant'].value_counts()[3]
DaTaFrAmE.at[0,'-3']=DaTaFrAmE['octant'].value_counts()[-3]
DaTaFrAmE.at[0,'4']=DaTaFrAmE['octant'].value_counts()[4]
DaTaFrAmE.at[0,'-4']=DaTaFrAmE['octant'].value_counts()[-4]

#Taking input from user.
mod=5000
DaTaFrAmE.at[1,'Octant ID']=mod

total_len=len(DaTaFrAmE['octant'])
temp1=0

#Using loop to split the data into different ranges.
while(total_len>0):
    temp2=mod

    #for last range.
    if total_len<mod:
        mod=total_len
    i=temp1*temp2
    j=temp1*temp2+mod-1

    #Here we will insert range and the corresponding data.
    DaTaFrAmE.at[temp1+2,'Octant ID']=str(i)+'-'+str(j)

    #Counting the octant, creating a new dataframe for a choosen range.
    DaTaFrAmE1=DaTaFrAmE.loc[i:j]

    #Inserting values into cell after checking the octant for a particular range.
    DaTaFrAmE.at[temp1+2,'1']=DaTaFrAmE1['octant'].value_counts()[1]
    DaTaFrAmE.at[temp1+2,'-1']=DaTaFrAmE1['octant'].value_counts()[-1]
    DaTaFrAmE.at[temp1+2,'2']=DaTaFrAmE1['octant'].value_counts()[2]
    DaTaFrAmE.at[temp1+2,'-2']=DaTaFrAmE1['octant'].value_counts()[-2]
    DaTaFrAmE.at[temp1+2,'3']=DaTaFrAmE1['octant'].value_counts()[3]
    DaTaFrAmE.at[temp1+2,'-3']=DaTaFrAmE1['octant'].value_counts()[-3]
    DaTaFrAmE.at[temp1+2,'4']=DaTaFrAmE1['octant'].value_counts()[4]
    DaTaFrAmE.at[temp1+2,'-4']=DaTaFrAmE1['octant'].value_counts()[-4]

    temp1=temp1+1
    total_len=total_len-mod