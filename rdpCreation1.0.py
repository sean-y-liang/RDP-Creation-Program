"""
Authored by Sean Liang on 6/28/2022.
This program mass generates RDP files for users given an Excel spreadsheet containing a list of host names. 
Tested using DNS spreadsheet exported from Cisco Meraki software, Windows 10 and 11 Remote Desktop Connection, and Python 3.11.
The path of the Excel file, the column name that contains the host names in the Excel file, the local domain name, the gateway hostname and the path for 
rdp file creation are taken as user input by a simple GUI.
Created RDP files are given the filename of their respective host name.
Requires installation of 'wheel', 'Pandas', 'openpyxl', 'Tkinter', and 'PySimpleGUI' libraries: 'pip install wheel pandas openpyxl tk pysimplegui'.
These libraries are already packaged with the executable.
"""

#Library imports.
import os
from sys import exit
import pandas as pd
import PySimpleGUI as sg
import tkinter as tk
from tkinter import simpledialog
import ctypes
import platform

#Configure GUI theme.
sg.theme('Default1')

#Make GUI match monitor resolution.
def make_dpi_aware():
    if int(platform.release()) >= 8:
        ctypes.windll.shcore.SetProcessDpiAwareness(True)

#Presents a UI for user input.
def input_Excel_file():
    #File explorer prompt to browse for Excel file.
    layout = [
        [sg.Text("Input Excel File:"), sg.Input(), sg.FileBrowse('Browse')],
        [sg.OK(), sg.Cancel()],
    ]
    window = sg.Window('RDP Creation 3.5', layout)
    event, values = window.read()
    if event is None or event == 'Cancel':
        window.close()
        exit(0)
    elif event == 'OK':   
        global dir_path
        dir_path = values['Browse'] 
        #Check that the field is not empty.
        if (dir_path == ""):
            window.close()
            input_Excel_file()
        window.close()

def input_column_name():
    #Define the input dialog for column name. Ex. 'HOSTNAME'
    global col_name
    col_name = simpledialog.askstring(title="RDP Creation 3.5",
    prompt="Input the title of the Excel column containing the hostnames: (ex. 'HOSTNAME')")
    #Remove whitespace from input string.
    col_name = col_name.rstrip()    
    #Check that the field is not empty.
    if (col_name == ""):
        input_column_name()  

def input_domain_name():
    #Define the input dialog for the domain name. Ex. '.CompanyX.local'
    global domain_name
    domain_name = simpledialog.askstring(title="RDP Creation 3.5",
    prompt="Input the domain name: (ex. '.CompanyX.local')")
    #Remove whitespace from input string.
    domain_name = domain_name.rstrip()
    #Check that the field is not empty.
    if (domain_name == ""):
        input_domain_name()

def input_gateway_hostname():
    #Define the input dialog for the gateway hostname. Ex. 'remote.CompanyX.com'
    global gateway_hostname
    gateway_hostname = simpledialog.askstring(title="RDP Creation 3.5",
    prompt="Input the gateway host name: (ex. 'remote.CompanyX.com'")
    #Remove whitespace from input string.
    gateway_hostname = gateway_hostname.rstrip()    
    #Check that the field is not empty.
    if (gateway_hostname == ""):
        input_gateway_hostname()

def creation_path():
    #File explorer prompt to declare the path of rdp file creation.
    layout = [
        [sg.Text("RDP Output Directory:"), sg.Input(), sg.FolderBrowse('Browse')],
        [sg.OK(), sg.Cancel()],
    ]
    window = sg.Window('RDP Creation 3.5', layout)
    event, values = window.read()
    if event is None or event == 'Cancel':
        window.close()
        exit(0)
    elif event == 'OK':   
        global write_path
        write_path = values['Browse']+'\\'
        #Check that the field is not empty.
        if (values['Browse'] == ""):
            window.close()
            creation_path()
        window.close()

#Feeds data into a Python list from an Excel spreadsheet.
def read_data():
    df = pd.read_excel(dir_path) 

    #Search the specified Excel column name for user host names.
    list_of_host_names = list(df[col_name])
    return list_of_host_names

#Converts the host names in a list to fqdm.
def fqdm_conversion(list_of_host_names):
    list_of_fqdm = [i + domain_name for i in list_of_host_names]
    return (list_of_fqdm)

#Creates a batch of rdp files using the previously entered data. Default values for rdp settings can be modified below.
def create_rdp_files(list_of_fqdm):
    for fqdm in list_of_fqdm:
        rdp_text = """screen mode id:i:2
use multimon:i:0
desktopwidth:i:1920
desktopheight:i:1080
session bpp:i:32
winposstr:s:0,3,0,0,800,600
compression:i:1
keyboardhook:i:2
audiocapturemode:i:0
videoplaybackmode:i:1
connection type:i:7
networkautodetect:i:1
bandwidthautodetect:i:1
displayconnectionbar:i:1
enableworkspacereconnect:i:0
disable wallpaper:i:0
allow font smoothing:i:0
allow desktop composition:i:0
disable full window drag:i:1
disable menu anims:i:1
disable themes:i:0
disable cursor setting:i:0
bitmapcachepersistenable:i:1
full address:s:"""+fqdm+"""
audiomode:i:0
redirectprinters:i:1
redirectcomports:i:0
redirectsmartcards:i:1
redirectclipboard:i:1
redirectposdevices:i:0
drivestoredirect:s:
autoreconnection enabled:i:1
authentication level:i:2
prompt for credentials:i:1
negotiate security layer:i:1
remoteapplicationmode:i:0
alternate shell:s:
shell working directory:s:
gatewayhostname:s:"""+gateway_hostname+"""
gatewayusagemethod:i:2
gatewaycredentialssource:i:4
gatewayprofileusagemethod:i:1
promptcredentialonce:i:1
gatewaybrokeringtype:i:0
use redirection server name:i:0
rdgiskdcproxy:i:0
kdcproxyname:s:
"""   
        with open(write_path+fqdm.removesuffix(domain_name)+'.rdp', 'w') as f:
            f.write(rdp_text)

    #Popup declaring whether the program succeeded or failed.
    if os.listdir(write_path):
        sg.Popup('Operation completed successfully.')
        exit(0)
    else:
        sg.Popup('Operation failed.')
        exit(1)

#Run the program.
make_dpi_aware()
input_Excel_file()
input_column_name()
input_domain_name()
input_gateway_hostname()
creation_path()
create_rdp_files(fqdm_conversion(read_data()))