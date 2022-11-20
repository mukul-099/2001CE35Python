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
    
    #Using regex.
    
    #These are some variables which will be storing the data which relates to their name.
    over=re.compile(r'(\d\d?\.\d)')
    zero=re.compile(r'no run')
    no_ball=re.compile(r', (no ball),')
    wide=re.compile(r', wide,')
    wide2=re.compile(r', 2 wides,')
    wide3=re.compile(r', 3 wides,')
    single=re.compile(r', 1 run,')
    six=re.compile(r', SIX,')
    four=re.compile(r', FOUR,')
    byes=re.compile(r', (byes),')
    lbyes=re.compile(r', (leg byes),')
    double=re.compile(r', 2 runs,')
    triple=re.compile(r', 3 runs,')
    out=re.compile(r', out')
    player=re.compile(r'(\d\d?\.\d) (\w+) (to|\w+ to) (\w+)( \w+)?,')
    caught=re.compile(r', out Caught by (\w+)')
    lbw=re.compile(r', out Lbw!!')
    bowled=re.compile(r', out Bowled!!')
    run_out=re.compile(r'Run Out!! ')
    runs=0
    wickets=0
    nb=0
    nlb=0
    nw=0
    nnb=0
    ppr=0
    
    #This code dynamically works on each line.
    while True:
        t = inn1.readline()
        if not t:
            break
        crun = 0
        cex = 0
        ov = over.findall(t)
        if float(ov[0]) == int(float(ov[0]))+0.6:
            ov[0] = str(int(float(ov[0]))+1)
        if float(ov[0]) == int(float(ov[0]))+0.1:
            orun = runs
        nnnb = no_ball.findall(t)
        wbb = wide.findall(t)
        sr = single.findall(t)
        zr = zero.findall(t)
        by = byes.findall(t)
        lby = lbyes.findall(t)
        pl = player.finditer(t)
        w2 = wide2.findall(t)
        w3 = wide3.findall(t)
        for i in pl:
            cbow = i.group(2)
            cbat = i.group(4)
        r = 1
        for col in SHEET1.iter_cols(min_col=1, max_col=1):
            for cell in col:
                if cell.value == bat_pl[cbat]:
                    crow = cell.row
                    r = 0
                    break
        if r:
            SHEET1.append([bat_pl[cbat], 'not out', '', '', '', 0, 0, 0, 0, 0])
            crow = SHEET1.max_row

        s = 1
        for col in SHEET2.iter_cols(min_col=1, max_col=1):
            for cell in col:
                if cell.value == bow_pl[cbow]:
                    cbrow = cell.row
                    s = 0
                    break
        if s:
            SHEET2.append([bow_pl[cbow], '', '', 0, 0, 0, 0, 0, 0, 0])
            cbrow = SHEET2.max_row
