This program mass generates RDP files for users given an Excel spreadsheet containing a list of host names. 
Tested using DNS spreadsheet exported from Cisco Meraki software, Windows 10 and 11 Remote Desktop Connection, and Python 3.11.
The path of the Excel file, the column name that contains the host names in the Excel file, the local domain name, the gateway hostname and the path for 
rdp file creation are taken as user input by a simple GUI.
Created RDP files are given the filename of their respective host name.
Requires installation of 'wheel', 'Pandas', 'openpyxl', 'Tkinter', and 'PySimpleGUI' libraries: 'pip install wheel pandas openpyxl tk pysimplegui'.
These libraries are packaged with the executable.
