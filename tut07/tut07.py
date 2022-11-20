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


def count_longest_subsequence_freq_func(longest, OUTsheet, total_count):
    # Dictionary to store consecutive sequence count
    count = {}

    # Dictionary to store frequency count
    frequency = {}

    # Dictionary to store time range
    timeRange = {}

    for label in OCT_SIGN:
        timeRange[label] = []

    # Initialing dictionary to 0 for all labels
    RST_CNT(count)
    RST_CNT(frequency)

    # Variable to check last value
    last = -10

    # Iterating complete excel sheet
    for i in range(0, total_count):
        currRow = i+3
        try:
            curr = int(OUTsheet.cell(column=11, row=currRow).value)

            # Comparing current and last value
            if (curr == last):
                count[curr] += 1
            else:
                count[curr] = 1
                RST_CNT_except(count, curr)

            # Updating frequency
            if (count[curr] == longest[curr]):
                frequency[curr] += 1

                # Counting STRTing and ending time of longest subsequence
                end = float(OUTsheet.cell(row=currRow, column=1).value)
                STRT = 100*end - longest[curr]+1
                STRT /= 100

                # Inserting time interval into map
                timeRange[curr].append([STRT, end])

                # Resetting count dictionary
                RST_CNT(count)
            else:
                RST_CNT_except(count, curr)
        except FileNotFoundError:
            print("File not found!!")
            exit()
        except ValueError:
            print("File content is invalid!!")
            exit()

        # Updating 'last' variable
        last = curr

    # Setting frequency table into sheet
    SET_FREQ(longest, frequency, OUTsheet)

    # Setting time range for longest subsequence
    longest_subsequence_time(longest, frequency, timeRange, OUTsheet)

# Method to set frequency count to sheet


def find_longest_subsequence(OUTsheet, total_count):
	# Dictionary to store consecutive sequence count
    count = {}

    # Dictionary to store longest count
    longest = {}

    # Initialing dictionary to 0 for all labels
    RST_CNT(count)
    RST_CNT(longest)

    # Variable to check last value
    last = -10

    # Iterating complete excel sheet
    for i in range(0, total_count):
        currRow = i+3
        try:
            curr = int(OUTsheet.cell(column=11, row=currRow).value)

            # Comparing current and last value
            if (curr == last):
                count[curr] += 1
                longest[curr] = max(longest[curr], count[curr])
                RST_CNT_except(count, curr)
            else:
                count[curr] = 1
                longest[curr] = max(longest[curr], count[curr])
                RST_CNT_except(count, curr)
        except FileNotFoundError:
            print("File not found!!")
            exit()

        # Updating "last" variable
        last = curr

    # Method to Count longest subsequence frequency
    count_longest_subsequence_freq_func(longest, OUTsheet, total_count)


def transition_count_func(row, transition_count, OUTsheet):
    # Setting hard coded inputs
    try:
        OUTsheet.cell(row=row, column=36).value = "To"
        OUTsheet.cell(row=row+1, column=35).value = "Octant #"
        OUTsheet.cell(row=row+2, column=34).value = "From"

        for i in range(35, 44):
            for j in range(row+1, row+1+9):
                OUTsheet.cell(row=j, column=i).border = BLC_border

    except FileNotFoundError:
        print("Output file not found!!")
        exit()
    except ValueError:
        print("Row or column values must be at least 1 ")
        exit()

    # Setting Labels
    for i, label in enumerate(OCT_SIGN):
        try:
            OUTsheet.cell(row=row+1, column=i+36).value = label
            OUTsheet.cell(row=row+i+2, column=35).value = label
        except FileNotFoundError:
            print("Output file not found!!")
            exit()
        except ValueError:
            print("Row or column values must be at least 1 ")
            exit()

    # Setting data
    for i, l1 in enumerate(OCT_SIGN):
        maxi = -1

        for j, l2 in enumerate(OCT_SIGN):
            val = transition_count[str(l1)+str(l2)]
            maxi = max(maxi, val)

        for j, l2 in enumerate(OCT_SIGN):
            try:
                OUTsheet.cell(row=row+i+2, column=36 +
                              j).value = transition_count[str(l1)+str(l2)]
                if transition_count[str(l1)+str(l2)] == maxi:
                    maxi = -1
                    OUTsheet.cell(row=row+i+2, column=36+j).fill = YELLOW_bg
            except FileNotFoundError:
                print("Output file not found!!")
                exit()
            except ValueError:
                print("Row or column values must be at least 1 ")
                exit()


def set_mod_overall_transition_count(OUTsheet, mod, total_count):
	# Counting partitions w.r.t. mod
    try:
        totalPartition = total_count//mod
    except ZeroDivisionError:
        print("Mod can't have 0 value")
        exit()

    # Checking mod value range
    if (mod < 0):
        raise Exception("Mod value should be in range of 1-30000")

    if (total_count % mod != 0):
        totalPartition += 1

    # Initializing row STRT for data filling
    rowSTRT = 16

    # Iterating all partitions
    for i in range(0, totalPartition):
        # Initializing STRT and end values
        STRT = i*mod
        end = min((i+1)*mod-1, total_count-1)

        # Setting STRT-end values
        try:
            OUTsheet.cell(column=35, row=rowSTRT-1 + 13 *
                          i).value = "Mod Transition Count"
            OUTsheet.cell(column=35, row=rowSTRT + 13 *
                          i).value = str(STRT) + "-" + str(end)
        except FileNotFoundError:
            print("Output file not found!!")
            exit()
        except ValueError:
            print("Row or column values must be at least 1 ")
            exit()

        # Initializing empty dictionary
        transCount = {}
        for a in range(1, 5):
            for b in range(1, 5):
                transCount[str(a)+str(b)] = 0
                transCount[str(a)+str(-b)] = 0
                transCount[str(-a)+str(b)] = 0
                transCount[str(-a)+str(-b)] = 0

        # Counting transition for range [STRT, end)
        for a in range(STRT, end+1):
            try:
                curr = OUTsheet.cell(column=11, row=a+3).value
                next = OUTsheet.cell(column=11, row=a+4).value
            except FileNotFoundError:
                print("Output file not found!!")
                exit()
            except ValueError:
                print("Row or column values must be at least 1 ")
                exit()

            # Incrementing count for within range value
            if (next != None):
                transCount[str(curr) + str(next)] += 1

        # Setting transition counts
        transition_count_func(rowSTRT + 13*i, transCount, OUTsheet)
