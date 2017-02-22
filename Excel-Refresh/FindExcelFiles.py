import glob, os
from ExcelRefresh import ExcelRefresh

path = "<complete file path to the folder>" # Use / instead of \

os.chdir(path)

i = 0
for file in glob.glob("*.xlsx"):
        ExcelRefresh(path)
        i= i+1


