# ExcelRefresh.py

def ExcelRefresh (filename):
    
    import win32com.client
    import win32con
    import shutil
    import time
    import ctypes
    import os
    from pathlib import Path
    from pythoncom import com_error

    file = filename
    SourcePathName = '<complete file to path the folder>' #Use / instead of \
    if os.path.exists(SourcePathName+file):
        # Open Excel
        Application = win32com.client.Dispatch("Excel.Application")
        # Shows Excel. Helps with debugging
        Application.Visible = 1
        # Open Your Workbook
        try:
            Workbook = Application.Workbooks.Open(SourcePathName)
        except com_error as reason:
            print reason
            quit()
        # Refesh connections
        Workbook.RefreshAll()
        # Delays for 1 second for the query to run
        time.sleep(1) 
        # Save and Close the Workbook
        Workbook.Save()
        Application.Quit()

        # Message Box
        MB_OK = 0x0
        MB_OKCXL = 0x01
        result = ctypes.windll.user32.MessageBoxA(0, 'File Refreshed: ' + file +'\nContinue?', 'Refreshing Excel files...', MB_OK | MB_OKCXL)
        if result == win32con.IDCANCEL:
            quit()
        return;
        



