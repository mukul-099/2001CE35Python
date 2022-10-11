#Mukul 2001CE35

#Importing pandas.
import pandas as pAnDa_jI

from datetime import datetime
start_time = datetime.now()

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
DaTaFrAmE=pAnDa_jI.read_excel("input_octant_longest_subsequence.xlsx")

#Pre-Processing the data.
DaTaFrAmE.at[0,'U_AVG']=DaTaFrAmE['U'].mean()
DaTaFrAmE.at[0,'V_AVG']=DaTaFrAmE['V'].mean()
DaTaFrAmE.at[0,'W_AVG']=DaTaFrAmE['W'].mean()

DaTaFrAmE['U-U_AVG']=DaTaFrAmE['U']-DaTaFrAmE.at[0,'U_AVG']
DaTaFrAmE['V-V_AVG']=DaTaFrAmE['V']-DaTaFrAmE.at[0,'V_AVG']
DaTaFrAmE['W-W_AVG']=DaTaFrAmE['W']-DaTaFrAmE.at[0,'W_AVG']

#Using "oCtAnT" func. for categorizing the data using .apply function.
DaTaFrAmE['Octant']=DaTaFrAmE.apply(lambda i: oCtAnT(i['U-U_AVG'],i['V-V_AVG'],i['W-W_AVG']),axis=1)

#Leaving empty column
DaTaFrAmE[''] = ''

#Defining columns for subsequnce.
DaTaFrAmE['Octant Num'] = ''
DaTaFrAmE['Longest Subsequence Length'] = ''
DaTaFrAmE['Count'] = ''

#Making a list of all the octants.
OcTaNt_NuM = [1, -1, 2, -2, 3, -3, 4, -4]

#Creating temprory variable(will be used in finding subsequence)
temp1 = 0

#Leaving empty column
DaTaFrAmE[''] = ''

#Defining columns for subsequnce.
DaTaFrAmE['Octant Num'] = ''
DaTaFrAmE['Longest Subsequence Length'] = ''
DaTaFrAmE['Count'] = ''

#Making a list of all the octants.
OcTaNt_NuM = [1, -1, 2, -2, 3, -3, 4, -4]

#Creating temprory variable(will be used in finding subsequence)
temp1 = 0

for i in OcTaNt_NuM:
    DaTaFrAmE.at[temp1, 'Octant Num'] = i
    cOuNt = 1
    temp2 = 1
    mAx_sUb_lEn = 0
    lst1 = DaTaFrAmE['Octant'].tolist()
    for j in range(len(lst1)-1):
        if (i == lst1[j] and i == lst1[j+1]):
            temp2 = temp2+1
        else:
            if (mAx_sUb_lEn == temp2):
                cOuNt = cOuNt+1
            elif (mAx_sUb_lEn < temp2):
                cOuNt = 1
            mAx_sUb_lEn = max(mAx_sUb_lEn, temp2)
            temp2 = 1
    DaTaFrAmE.at[temp1, 'Longest Subsequence Length'] = mAx_sUb_lEn
    DaTaFrAmE.at[temp1, 'Count'] = cOuNt
    temp1 = temp1+1

#Now storing this dataframe to excel file.
DaTaFrAmE.to_excel('output_octant_longest_subsequence.xlsx')

#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))



