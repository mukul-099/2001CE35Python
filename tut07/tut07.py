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
