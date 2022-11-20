#Mukul 2001CE35

#Calling all required libraries.
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment,Border,Side
import numpy as NUM_py
import os
import re
from datetime import datetime
os.system('cls')
start_time=datetime.now()

#This function will work only for one inning.
#We will call it again for the second inning.
def INNING(inn1,bat_pl,bow_pl,s): 
    innbat=Workbook()
    innfow=Workbook()
    innbow=Workbook()
    SHEET1=innbat.active
    SHEET2=innbow.active
    SHEET3=innfow.active
    SHEET1.column_dimensions['A'].width=25

    #Creating column headings for scorecard and index.
    SHEET1['A1']=s+' Innings'
    SHEET1['I1']='0-0'
    SHEET1['J1']='0 overs'
    SHEET1['A2']='Batter'
    SHEET1['F2']='R'
    SHEET1['G2']='B'
    SHEET1['H2']='4s'
    SHEET1['I2']='6s'
    SHEET1['J2']='SR'
    SHEET2['A1']='Bowler'
    SHEET2['D1']='O'
    SHEET2['E1']='M'
    SHEET2['F1']='R'
    SHEET2['G1']='W'
    SHEET2['H1']='NB'
    SHEET2['I1']='WD'
    SHEET2['J1']='ECO'
    SHEET3['A1']='Fall of wickets'