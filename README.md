# NameSync

Windows 10 application to aid file/folder syncing.

You might have come across situations where in you had a backup of your files/folders in a computer (Sink) which you do not use much, and without network perhaps. You might want to compare whether all the files/folders in your main computer (Source) that you use more is in backup Sink computer.
Well, "NameSync" let's you do just that. You can think of it as a primitive "git". It checks whether all files/folders in Source are present in Sink. Beware, NameSync checks only the names and not the modifications that the files/folders might have. Follow the below steps to give it a try.

1. Copy and run "Sink.exe" inside the backup folder of Sink computer. This will generate "SyncInfo.xlsx".
2. Copy "SyncInfo.xlsx" from Sink computer to folder corresponding to backup folder in Source computer.
3. Run "Source.exe" in the folder.
4. Carefully read "SyncInfo.xlsx" and manually copy disparity files/folders from Source computer to Sink computer.

The versions of softwares used are as given below.

python version:   Python 3.8.3
pip3 version:     pip 20.1.1
openpyxl version: 3.0.4