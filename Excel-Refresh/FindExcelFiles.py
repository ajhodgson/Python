import glob, os
from ExcelRefresh import ExcelRefresh

path = "<complete file path to the folder>"
print("\nPath:  " + path + "\n")
os.chdir(path)

i = 0
for file in glob.glob("*.xlsx"):
        ExcelRefresh(file, path)
        i= i+1
quit();


