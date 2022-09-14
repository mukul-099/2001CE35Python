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
