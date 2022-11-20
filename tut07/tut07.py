#Mukul
#2001CE35

#CAlling all required lib.

from platform import python_version
from openpyxl.styles import Color, PatternFill, Font, Border, Side
import glob
import openpyxl
import os
from datetime import datetime
STRT_time = datetime.now()
os.system("cls")


ver = python_version()

if ver == "3.8.10":
	print("Correct Version Installed")
else:
	print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")

##Read all the excel files in a batch format from the input/ folder. Only xlsx to be allowed
##Save all the excel files in a the output/ folder. Only xlsx to be allowed
## output filename = input_filename[_octant_analysis_mod_5000].xlsx , ie, append _octant_analysis_mod_5000 to the original filename.

OCT_SIGN = [1, -1, 2, -2, 3, -3, 4, -4]
OCT_NAME_id_mapping = {1: "Internal outward interaction", -1: "External outward interaction", 2: "External Ejection", -
                       2: "Internal Ejection", 3: "External inward interaction", -3: "Internal inward interaction", 4: "Internal sweep", -4: "External sweep"}
YELLOW = "00FFFF00"
YELLOW_bg = PatternFill(
	start_color=YELLOW, end_color=YELLOW, fill_type='solid')
BLC = "00000000"
DoUbLe = Side(border_style="thin", color=BLC)
BLC_border = Border(top=DoUbLe, left=DoUbLe, right=DoUbLe, bottom=DoUbLe)

#Code


def RST_CNT(count):
    for item in OCT_SIGN:
        count[item] = 0

# Method to initialise dictionary with 0 for "OCT_SIGN" except 'left'


def RST_CNT_except(count, left):
    for item in OCT_SIGN:
        if (item != left):
            count[item] = 0


def SET_FREQ(longest, frequency, OUTsheet):
    # Iterating "OCT_SIGN" and updating sheet
    for i in range(9):
        for j in range(3):
            OUTsheet.cell(row=3+i, column=45+j).border = BLC_border

    OUTsheet.cell(row=3, column=45).value = "Octant ##"
    OUTsheet.cell(row=3, column=46).value = "Longest Subsquence Length"
    OUTsheet.cell(row=3, column=47).value = "Count"

    for i, label in enumerate(OCT_SIGN):
        currRow = i+3
        try:
            OUTsheet.cell(row=currRow+1, column=45).value = label
            OUTsheet.cell(column=46, row=currRow+1).value = longest[label]
            OUTsheet.cell(column=47, row=currRow+1).value = frequency[label]
        except FileNotFoundError:
            print("File not found!!")
            exit()

# Method to set time range for longest subsequence


def longest_subsequence_time(longest, frequency, timeRange, OUTsheet):
    # Naming columns number
    lengthCol = 50
    freqCol = 51

    # Initial row, just after the header row
    row = 4

    OUTsheet.cell(row=3, column=49).value = "Octant ###"
    OUTsheet.cell(row=3, column=50).value = "Longest Subsquence Length"
    OUTsheet.cell(row=3, column=51).value = "Count"

    # Iterating all octants
    for octant in OCT_SIGN:
        try:
            # Setting octant's longest subsequence and frequency data
            OUTsheet.cell(column=49, row=row).value = octant
            OUTsheet.cell(column=lengthCol, row=row).value = longest[octant]
            OUTsheet.cell(column=freqCol, row=row).value = frequency[octant]
        except FileNotFoundError:
            print("File not found!!")
            exit()

        row += 1

        try:
            # Setting default labels
            OUTsheet.cell(column=49, row=row).value = "Time"
            OUTsheet.cell(column=lengthCol, row=row).value = "From"
            OUTsheet.cell(column=freqCol, row=row).value = "To"
        except FileNotFoundError:
            print("File not found!!")
            exit()

        row += 1

        # Iterating time range values for each octants
        for timeData in timeRange[octant]:
            try:
                # Setting time interval value
                OUTsheet.cell(row=row, column=lengthCol).value = timeData[0]
                OUTsheet.cell(row=row, column=freqCol).value = timeData[1]
            except FileNotFoundError:
                print("File not found!!")
                exit()
            row += 1

    for i in range(3, row):
        for j in range(49, 52):
            OUTsheet.cell(row=i, column=j).border = BLC_border
