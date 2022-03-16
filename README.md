# NameSync

Windows 11 64-bit application to aid file/folder syncing.

This software compares names of files/folders in two seperate locations in same computer or different computers. Then it comes up with an Excel file ("SyncInfo.xlsx") that informs the files/folders to be copied between the locations/computers to make both of them in sync.

Beware, this software compares only the names and not their contents. So, files with same name but different modifications will not be detected.

Steps to run the software

1. Run "Comp_1.exe" in first computer inside the folder/drive to be synced thus creating "SyncInfo.xlsx"
2. Copy "SyncInfo.xlsx" to the concerned folder/drive in second computer
3. Run "Comp_2.exe" in concerned folder/drive in second computer thus modifying "SyncInfo.xlsx"
4. Carefully read "SyncInfo.xlsx" and manually copy the discrepancy files/folders between the computers

The versions of softwares used, are as given below.

python version: Python 3.10.2, pip3 version: pip 22.0.4, openpyxl version: 3.0.9
