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
    
#Creating a function to find transition count.


def cOuNt_TrAnSiTiOnS(Dataframe, p, q):
    r = 0
    for t in range(len(Dataframe)-1):
        if Dataframe.at[t, 'octant'] == p and Dataframe.at[t+1, 'octant'] == q:
            r = r+1
    return r


sp = int(len(DaTaFrAmE)/mod)
DaTaFrAmE.at[sp+6, 'Octant ID'] = 'Overall Transition Count'
DaTaFrAmE.at[sp+7, 'Octant ID'] = 'To'
DaTaFrAmE.at[sp+8, '1'] = 1
DaTaFrAmE.at[sp+8, '-1'] = -1
DaTaFrAmE.at[sp+8, '2'] = 2
DaTaFrAmE.at[sp+8, '-2'] = -2
DaTaFrAmE.at[sp+8, '3'] = 3
DaTaFrAmE.at[sp+8, '-3'] = -3
DaTaFrAmE.at[sp+8, '4'] = 4
DaTaFrAmE.at[sp+8, '-4'] = -4

DaTaFrAmE.at[sp+9, ''] = 'From'
DaTaFrAmE.at[sp+8, 'Octant ID'] = 'Count'
DaTaFrAmE.at[sp+9, 'Octant ID'] = -4
DaTaFrAmE.at[sp+10, 'Octant ID'] = -3
DaTaFrAmE.at[sp+11, 'Octant ID'] = -2
DaTaFrAmE.at[sp+12, 'Octant ID'] = -1
DaTaFrAmE.at[sp+13, 'Octant ID'] = 1
DaTaFrAmE.at[sp+14, 'Octant ID'] = 2
DaTaFrAmE.at[sp+15, 'Octant ID'] = 3
DaTaFrAmE.at[sp+16, 'Octant ID'] = 4

#Now we'll calculate overall transition count by calling cOuNt_TrAnSiTiOnS.
for i in range(int(len(DaTaFrAmE)/mod)+9, int(len(DaTaFrAmE)/mod)+13):
    for j in range(-4, 5):
        DaTaFrAmE.at[i, str(j)] = cOuNt_TrAnSiTiOnS(
            DaTaFrAmE, i-int(len(DaTaFrAmE)/mod)-13, j)
for i in range(int(len(DaTaFrAmE)/mod)+13, int(len(DaTaFrAmE)/mod)+17):
    for j in range(-4, 5):
        DaTaFrAmE.at[i, str(j)] = cOuNt_TrAnSiTiOnS(
            DaTaFrAmE, i-int(len(DaTaFrAmE)/mod)-12, j)

total_len = len(DaTaFrAmE['octant'])


#Defining a new function for counting mod transitions.
temp4 = 1


def CoUnt_mod_TrAnSiTiOnS(Dataframe, mod, p, q):
    r = 0
    if (mod*temp4-1 < len(Dataframe)):
        for t in range(mod*(temp4-1), mod*temp4-1):
            if Dataframe.at[t, 'octant'] == p and Dataframe.at[t+1, 'octant'] == q:
                r = r+1
    else:
        for t in range(mod*(temp4-1), len(Dataframe)-1):
            if Dataframe.at[t, 'octant'] == p and Dataframe.at[t+1, 'octant'] == q:
                r = r+1
    return r


mod = 5000
