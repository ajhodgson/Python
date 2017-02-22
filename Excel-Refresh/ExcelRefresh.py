# ExcelRefresh.py
import win32com.client
import win32con
import shutil
import time
import ctypes
import os
from pathlib import Path
from pythoncom import com_error

def ExcelRefresh (filename):
    file = filename
    SourcePathName = '<complete file to path the folder>' #Use / instead of \
    if os.path.exists(SourcePathName+file):
        # Open Excel
        Application = win32com.client.Dispatch("Excel.Application")
        # Makes Excel visible (1 = visible)
        Application.Visible = 0
        # Open Your Workbook
        try:
            Workbook = Application.Workbooks.Open(SourcePathName+file)
        except com_error as reason:
            print reason
            quit()
        # Refesh connections
        Workbook.RefreshAll()
        # Delays for 1 second for the query to run
        time.sleep(1) 
        # Save and Close
        Workbook.Save()
        Application.Quit()
        return;
    else:
        quit();
        



