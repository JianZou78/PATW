# -*- coding: utf-8 -*-
"""
Created on Fri April  12 19:20:25 2025

@author: jianzou
"""

import pygetwindow as gw
import pyautogui
import time
import subprocess
import tkinter as tk
from tkinter import filedialog
import os
import shutil
import stat
import wexpect
from datetime import datetime
import pyperclip
import win32gui
import csv
import paramiko
from scp import SCPClient
import re  
import requests
import numpy as np



def get_input(prompt, default_value):
    user_input = input(f"{prompt} (default: {default_value}): ")
    return user_input if user_input else default_value
PATW_version = r"Release1 Version: April 20, 2025"    
print("\n\n---------------------------------PATW_Python_Auto_Teams_MTRw MPAT------------------------------------------------------\n")
print('Tool for development testing only. Provided "as is" with no guarantee of compatibility with your test environment.\n\n')
print('Prerequisite\n')
print('Install MPAT on the x64 Windows PC. \n')
print('Ensure Windows Teams is installed and can call DUT MTRw Teams \n')
print('Install TightVNC on the Windows PC and MTRw, Windows PC can TightVNC to the DUT MTRw \n')
print('Windows PC can ssh connect to DUT MTRw Teams \n\n')
print(f"------------------------------------------{PATW_version}----------------------------------------------------\n\n")



testMethod = get_input('TightVNC to accept/end call, run Capture_MousePosition firstly. MPAT method for example: upload\n','upload')

call_to_address = get_input('Please input email address to be called, for example: sox_dut@MODERNCOMMS784471.onmicrosoft.com\n','sox_dut@MODERNCOMMS784471.onmicrosoft.com')

call_audio_video = get_input('audio call or video call , for example: video\n','video')

call_to_win_name = get_input('Please input call to window name keywords, for example: sox_dut\n','sox_dut')

inCall_win_name = get_input('Please input incall window name keywords, for example: Microsoft Teams meeting\n','Microsoft Teams meeting')

TightVNC_win_name = get_input('Please input TightVNC window title, for example: desktop-rfsncel\n','desktop-rfsncel')

call_Duration =int( get_input('Please input call duration expected in seconds , for example:10\n','10'))

dut_ipaddr = get_input('Please input the ip address of the DUT, for example: 10.168.4.32\n','10.168.4.32')

MTRw_usr = get_input('Please input the administrator to login MTRw PC, for example: admin\n','admin')

MTRw_pw = get_input('Please input the password to login MTRw PC, for example: sfb\n','sfb')

local_path = get_input('Plaease input the local path to save MTRw logs, for example: c:\MTRwLog\n',r'c:\MTRwLog')

vqe_enable = get_input('Enable VQE Recording on MTRw? Yes to input Y, No to input N, for example: Y\n',r'Y')

loopN = int(get_input('Please input the test loop, for example: 3\n','3'))


def extract_plots_from_mpat_url(url = 'https://logvisualizer.azurewebsites.net/audio/?id=audio-novalidinfo&prefix=83d843c9-97f4-4d6a-951d-06383cb59e41'):
    
    # Fetch the data from the URL
    #url = 'https://logvisualizer.azurewebsites.net/audio/?id=audio-novalidinfo&prefix=83d843c9-97f4-4d6a-951d-06383cb59e41'
    response = requests.get(url)
    
    # Print the response content for verification
    #print(response.text)
    
    # Extract the section of interest for Nearend Input (dbFS)
    section = re.search(r'Nearend Input \(dbFS\).*?keyList\.push\("Nearend Input \(channel 0\)"\)', response.text, re.DOTALL)
    null_Points=0
    
    nearendinput0 = []
    nearendoutput = []
    farendinput = []
    
    
    if section:
        section = section.group()
        # Extract the values pushed to valueList within the section
        value_lists = re.findall(r'valueList\.push\((.*?)\);', section, re.DOTALL)
        # Clean and convert the extracted values to a list of floats
        value_lists = [list(map(float, re.findall(r'-?\d+\.\d+', value_list))) for value_list in value_lists]
        nearendinput0 = value_lists
        for index in range(0,len(nearendinput0)):
            if nearendinput0[index] == []:
                nearendinput0[index] = [-125]
                null_Points =null_Points + 1
        
        nearendinput0_null = null_Points
        #print("\nnearendinput0\n")
        #print(nearendinput0)
    else:
        nearendinput0 = [-250]
        nearendinput0_null = 0
        
        print("No match found for Nearend Input (dbFS)")
    
    # Extract the section of interest for Nearend Output (dbFS)
    section = re.search(r'Nearend Output \(dbFS\).*?keyList\.push\("Nearend Output"\)', response.text, re.DOTALL)
    null_Points=0
    if section:
        section = section.group()
        # Extract the values pushed to valueList within the section
        value_lists = re.findall(r'valueList\.push\((.*?)\);', section, re.DOTALL)
        # Clean and convert the extracted values to a list of floats
        value_lists = [list(map(float, re.findall(r'-?\d+\.\d+', value_list))) for value_list in value_lists]
        nearendoutput = value_lists
        for index in range(0,len(nearendoutput)):
            if nearendoutput[index] == []:
                nearendoutput[index] = [-125]
                null_Points =null_Points + 1
        nearendoutput_null = null_Points
        
        #print("\nnearendoutput\n")
        #print(nearendoutput)
    else:
        nearendoutput = [-250]
        nearendoutput_null = 0
        print("No match found for Nearend Output (dbFS)")
    
    # Extract the section of interest for Farend Input (dbFS)
    section = re.search(r'Farend Input \(dbFS\).*?keyList\.push\("Farend Input"\)', response.text, re.DOTALL)
    null_Points=0
    if section:
        section = section.group()
        # Extract the values pushed to valueList within the section
        value_lists = re.findall(r'valueList\.push\((.*?)\);', section, re.DOTALL)
        # Clean and convert the extracted values to a list of floats
        value_lists = [list(map(float, re.findall(r'-?\d+\.\d+', value_list))) for value_list in value_lists]
        farendinput = value_lists
        for index in range(0,len(farendinput)):
            if farendinput[index] == []:
                farendinput[index] = [-125]
                null_Points =null_Points + 1
        
        farendinput_null = null_Points
        #print("\nfarendinput\n")
        #print(farendinput)
    else:
        farendinput = [-250]
        farendinput_null = 0
        print("No match found for Farend Input (dbFS)")
    
    
    
    
    
    # Calculate average levels
    avg_nearend_input = np.mean(nearendinput0)
    avg_nearend_output = np.mean(nearendoutput)
    avg_farend_input = np.mean(farendinput)
    
    # # Plot the data
    # plt.figure(figsize=(10, 5))
    # plt.plot(nearendinput0, label='Nearend Input')
    # plt.plot(nearendoutput, label='Nearend Output')
    # plt.axhline(y=avg_nearend_input, color='r', linestyle='--', label='Avg Nearend Input')
    # plt.axhline(y=avg_nearend_output, color='g', linestyle='--', label='Avg Nearend Output')
    # plt.xlabel('Time')
    # plt.ylabel('Level')
    # plt.title('Nearend Input vs Output Levels')
    # plt.legend()
    # plt.show()
    
    return nearendinput0,nearendoutput,farendinput,len(nearendinput0),len(nearendoutput),len(farendinput),nearendinput0_null,nearendoutput_null,farendinput_null,avg_nearend_input,avg_nearend_output,avg_farend_input




def create_ssh_client(server, port, user, password):
    client = paramiko.SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(server, port, user, password)
    return client

def transfer_folder_from_remote(ssh_client, local_path, remote_path):
    with SCPClient(ssh_client.get_transport()) as scp:
        scp.get(remote_path,local_path, recursive=True)

def transfer_logs_from_MTRw(remote_ip = dut_ipaddr,local_path = r'C:\recordings'):
    
    # Variables
    
    username = MTRw_usr
    password = MTRw_pw
    
   # Check if local_path exists
    if not os.path.exists(local_path):
        os.makedirs(local_path)
    else:
        # Empty the local_path folder
        for filename in os.listdir(local_path):
            file_path = os.path.join(local_path, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')
    
    # Transfer the file
    remote_path='C:/Users/Skype/AppData/Local/Packages/MSTeamsRooms_8wekyb3d8bbwe/LocalCache/Microsoft/MSTeams/Logs/sc-tfw/mediastack'
    #transfer_file(remote_ip, username, password, remote_path, local_path)
    ssh_client = create_ssh_client(dut_ipaddr, 22, username, password)
    transfer_folder_from_remote(ssh_client, local_path, remote_path)
    
    print(f"File transferred from {remote_ip}:{remote_path} to {local_path}\n")
    
    remote_path='C:/Users/Skype/AppData/Local/Packages/MSTeamsRooms_8wekyb3d8bbwe/LocalCache/Microsoft/MSTeams/Logs/sc-tfw/skylib'
    transfer_folder_from_remote(ssh_client, local_path, remote_path)
    
    print(f"File transferred from {remote_ip}:{remote_path} to {local_path}\n")
    
    remote_path='C:/Users/Skype/AppData/Local/Packages/MSTeamsRooms_8wekyb3d8bbwe/LocalCache/Microsoft/MSTeams/Logs/sc-tfw/SkypeRT/persistent.conf'
    transfer_folder_from_remote(ssh_client, os.path.join(local_path,'persistent.conf.txt'), remote_path)
    
    print(f"File transferred from {remote_ip}:{remote_path} to {local_path}\n")
    
    remote_path='C:/Users/Skype/AppData/Local/Packages/Microsoft.SkypeRoomSystem_8wekyb3d8bbwe/LocalState/Tracing/RigelAppSettings.txt'
    transfer_folder_from_remote(ssh_client, os.path.join(local_path,'RigelAppSettings.txt'), remote_path)
    
    print(f"File transferred from {remote_ip}:{remote_path} to {local_path}\n")
    
    remote_path='C:/recordings'
    transfer_folder_from_remote(ssh_client, local_path, remote_path)
    
    print(f"File transferred from {remote_ip}:{remote_path} to {local_path}\n")
    
    ssh_client.close()
    

def ensure_remote_folder_exists(ssh_client, remote_path):
    stdin, stdout, stderr = ssh_client.exec_command(f'mkdir -p {remote_path}')
    stdout.channel.recv_exit_status()  # Wait for the command to complete

def transfer_folder_to_remote(ssh_client, local_path, remote_path):
    ensure_remote_folder_exists(ssh_client, remote_path)
    with SCPClient(ssh_client.get_transport()) as scp:
        for root, dirs, files in os.walk(local_path):
            for file in files:
                local_file = os.path.join(root, file)
                remote_file = os.path.join(remote_path, os.path.relpath(local_file, local_path)).replace("\\", "/")
                print(f"Transferring {local_file} to {remote_file}")
                scp.put(local_file, remote_file)


def MTRw_EnableVQE(hostname, port, username, password):
    '''
    1. Transfer folder ControlVQE_MTRwT2_1_RevB from local_pc to MTRw PC c:/recordings
    '''
    # Initialize the SSH client
    ssh_client = create_ssh_client(hostname, port, username, password)
    
    # Define paths
    local_path = os.path.join(os.getcwd(), "ControlVQE_MTRwT2_1_RevB")
    remote_path = "c:/ControlVQE_MTRwT2_1_RevB"
    
    transfer_folder_to_remote(ssh_client, local_path, remote_path)
    
    
     
    # ssh_client = create_ssh_client(hostname, port, username, password)
    # remote_cmd_path = "c:\ControlVQE_MTRw_T2_1_RevB\VQEAPIRecordings_Enable_MTRwT2_1_RevB.cmd"
    # command = f"cmd /c {remote_cmd_path}"
    
    # # Execute the PowerShell command
    # # Execute the PowerShell command
    # stdin, stdout, stderr = ssh_client.exec_command(f'powershell.exe {"logoff console"}')
    # time.sleep(2)
    
      
    # print('click sign in in the MTRw PC\n')
    
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    
    # Original PowerShell command with potential non-breaking space
    powershell_command = r"PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^ .\PersistentConfigMTRwT2_1_RevB.ps1 -app_name Microsoft.SkypeRoomSystem -set -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings' -val '1'"
    # Remove non-breaking space (U+00A0) if present
    powershell_command = powershell_command.replace('\u00A0', ' ')
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    command = f'cd {remote_folder} && {powershell_command}'
    stdin, stdout, stderr = ssh_client.exec_command(command)
    
    time.sleep(1)

    # Original PowerShell command with potential non-breaking space
    powershell_command = r"PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^ .\PersistentConfigMTRwT2_1_RevB.ps1 -app_name Microsoft.SkypeRoomSystem -set -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName' -val 'c:\recordings\VqeApiRec*.dat'"
    # Remove non-breaking space (U+00A0) if present
    powershell_command = powershell_command.replace('\u00A0', ' ')
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    command = f'cd {remote_folder} && {powershell_command}'
    stdin, stdout, stderr = ssh_client.exec_command(command)
    
    

    
    ssh_client.close()
    

def MTRw_DisableVQE(hostname, port, username, password):
    '''
    1. Transfer folder ControlVQE_MTRwT2_1_RevB from local_pc to MTRw PC c:/recordings
    '''
        
    # Initialize the SSH client
    ssh_client = create_ssh_client(hostname, port, username, password)
    
    # Define paths
    local_path = os.path.join(os.getcwd(), "ControlVQE_MTRwT2_1_RevB")
    remote_path = "c:/ControlVQE_MTRwT2_1_RevB"
    
    transfer_folder_to_remote(ssh_client, local_path, remote_path)
    
    
     
    # ssh_client = create_ssh_client(hostname, port, username, password)
    # remote_cmd_path = "c:\ControlVQE_MTRw_T2_1_RevB\VQEAPIRecordings_Enable_MTRwT2_1_RevB.cmd"
    # command = f"cmd /c {remote_cmd_path}"
    
    # # Execute the PowerShell command
    # # Execute the PowerShell command
    # stdin, stdout, stderr = ssh_client.exec_command(f'powershell.exe {"logoff console"}')
    # time.sleep(2)
    
      
    # print('click sign in in the MTRw PC\n')
    
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    
    # Original PowerShell command with potential non-breaking space
    powershell_command = r"PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^ .\PersistentConfigMTRwT2_1_RevB.ps1 -app_name Microsoft.SkypeRoomSystem -del -domain 'MsrtcEcs' -key 'ADSP\EnableUnifiedVQEAPIRecordings'"
    # Remove non-breaking space (U+00A0) if present
    powershell_command = powershell_command.replace('\u00A0', ' ')
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    command = f'cd {remote_folder} && {powershell_command}'
    stdin, stdout, stderr = ssh_client.exec_command(command)
    
    time.sleep(1)

    # Original PowerShell command with potential non-breaking space
    powershell_command = r"PowerShell -NoProfile -ExecutionPolicy Bypass -Command ^ .\PersistentConfigMTRwT2_1_RevB.ps1 -app_name Microsoft.SkypeRoomSystem -del -domain 'MsrtcEcs' -key 'ADSP\UnifiedVQEAPIRecordingFileName'"
    # Remove non-breaking space (U+00A0) if present
    powershell_command = powershell_command.replace('\u00A0', ' ')
    remote_folder = "c:/ControlVQE_MTRwT2_1_RevB"
    command = f'cd {remote_folder} && {powershell_command}'
    stdin, stdout, stderr = ssh_client.exec_command(command)
    
    

    
    ssh_client.close()
    
    MTRw_logoff_cleanLog(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw)
    
    # login again
    
    while True:
        exists_TightVNC_win = window_exists_exact(TightVNC_win_name)
        
        if exists_TightVNC_win:
            window_list = gw.getWindowsWithTitle(TightVNC_win_name)
            TightVNC_window = window_list[0]
            time.sleep(3)
            #window_call.close()
            print(f"TightVNC Window with title '{TightVNC_win_name}' exists.")
            TightVNC_window.activate()
            TightVNC_window.maximize()
            break
        else:
            print(f"TightVNC Window with title '{TightVNC_win_name}' not exists.")
            time.sleep(2)
    
    # MTRw Sign in
    call_Pos = load_Call_Positions_from_csv(os.path.join(os.getcwd(),'MTRw_Positions.csv'))
    new_mouse_x = TightVNC_window.left + int(call_Pos[1][1])
    new_mouse_y = TightVNC_window.top + int(call_Pos[1][2])
    
    pyautogui.moveTo(new_mouse_x, new_mouse_y)
    time.sleep(1)
    pyautogui.click()
    
    # wait enough time for MTRw to launch
    time.sleep(30)
    
    
    
                


def window_exists_exact(title):
    def enum_windows_callback(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            window_text = win32gui.GetWindowText(hwnd)
            if title in window_text:
                titles.append(hwnd)

    titles = []
    win32gui.EnumWindows(enum_windows_callback, titles)
    return len(titles) > 0


def TightVNC_run(TightVNCExePath = r"C:\TightVNC-win64-v2.1"):
    
    currPath = os.getcwd()
    
    #TightVNCExePath = r"C:\TightVNC-win64-v2.1"
    os.chdir(TightVNCExePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(3)
    
    cmdline1 = r'.\TightVNC.exe -m 1920'
     
    
    commands = cmdline1;
     
    window_list = gw.getWindowsWithTitle('PowerShell')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        window.activate()
        time.sleep(1)
        pyautogui.typewrite(commands)
        pyautogui.press('enter')

        #os.chdir(currPath);
        time.sleep(2)
        window.close()
    else:
        print("No window with the title 'PowerShell' found.")
        
    os.chdir(currPath)
        
    # window_list = gw.getWindowsWithTitle(TightVNC_win_name)
    
    # # Check if any windows were found
    # if window_list:
    #     # Select the first window from the list
    #     window = window_list[0]

    #     window.close()
    # else:
    #     print(f"No window with the title {TightVNC_win_name} found.\n")
        

def select_Putty_folder():
    
    print (" select the Putty folder path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    Putty_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {Putty_path}")
    return Putty_path
    

def select_MPAT_folder():
    
    print (" select the MPAT folder path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    MPAT_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {MPAT_path}")
    return MPAT_path

def select_ResultSave_folder():
    
    print (" select the ResultSave path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    ResultSave_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {ResultSave_path}")
    return ResultSave_path
def find_window_containing_string(search_string):
    while True:
        # Get a list of all window titles
        windows = gw.getAllTitles()

        # Filter windows that contain the search string in their title
        matching_windows = [title for title in windows if search_string.lower() in title.lower()]

        if matching_windows:
            # Get the window object of the first matching window
            window = gw.getWindowsWithTitle(matching_windows[0])[0]
            print(f"Window ID: {window._hWnd}, Title: {window.title}")
            break

        # Wait for a short period before checking again
        time.sleep(1)
    return window

def open_cmd_and_get_window_id(directory):
    # Open the command prompt and change to the given directory
    subprocess.Popen(f'start cmd /K "cd /d {directory}"', shell=True)
    
    windowsList= find_window_containing_string(r'C:\WINDOWS\system32\cmd.exe')

    #windowsList = gw.getWindowsWithTitle('C:\WINDOWS\system32\cmd.exe')
    
    return windowsList
      





def auto_WinDesk_call_MTRA():
    
    start_time = time.time()
    
          
    # Get all windows
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Print all window titles
    for item in window_list:
        if str.lower('Microsoft Teams')in item.title.lower():
            print(f"Teams Window found in {item.title}\n")
            if call_to_win_name.lower() in item.title.lower():
                item.close()
            elif inCall_win_name.lower() in item.title.lower():
                item.close()
            else:
                print('...')
                #item.activate()
                
                
            
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams Main windows found\n")
        # Activate the window
        window.maximize()
    
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")
    
   
    #time.sleep(3)
    #window.activate()
    pyautogui.hotkey('ctrl', 'shift','N')
    
    time.sleep(5)

    time.sleep(1)

    pyautogui.typewrite(call_to_address)
    time.sleep(1)
 
    pyautogui.press('enter')
    # time.sleep(1)
    while True:
        exists_call_to_win = window_exists_exact(call_to_win_name)
        exists_in_call_win = window_exists_exact(inCall_win_name)
        
        if (exists_call_to_win or exists_in_call_win) :
            window_list = gw.getWindowsWithTitle(call_to_win_name)
            window_call = window_list[0]
            time.sleep(1)
            #window_call.close()
            print(f"Window with title '{call_to_win_name}' or '{inCall_win_name}' exists.")
            break
        else:
            print(f"Window with title '{call_to_win_name}' and '{inCall_win_name}' does not exist.")
            pyautogui.press('enter')
            time.sleep(2)
    
    pyautogui.press('enter')
    
    time.sleep(1)
    # Send Alt+Shift+V, start video call
    if call_audio_video.lower() in "video":
        pyautogui.hotkey('alt', 'shift', 'v')
    elif call_audio_video.lower() in "audio":
        pyautogui.hotkey('alt', 'shift', 'a')
    else:
        print('Please check input if "audio" or "video')
        
    
        
         
    
    time.sleep(3)
    
    #===============================================================================start Call==============================

    
    #===============================================================================End Call==============================
    time.sleep(call_Duration)  # call duration
    #===============================================================================End Call==============================

    
    time.sleep(2)
    

 
    
        
    # Print the output of the command
    #print(result.stdout)
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle(call_to_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{call_to_win_name}" ...\n')
            print(item.title)
            if call_to_win_name.lower() in item.title.lower():
                item.activate()
                
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
            
    else:
        print(f"No window with the title '{call_to_win_name}' found...\n")
     
    time.sleep(1)
    
    window_list = gw.getWindowsWithTitle(inCall_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{inCall_win_name}" ...\n')
            print(item.title)
            if inCall_win_name.lower() in item.title.lower():
                item.activate()
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
    else:
        print(f"No window with the title '{inCall_win_name}' found...\n")
    # exists_call_to_win = window_exists_exact(call_to_win_name)
    # exists_in_call_win = window_exists_exact(inCall_win_name)
    
  
    # if exists_call_to_win :
    #     window_list = gw.getWindowsWithTitle(call_to_win_name)
    #     # Select the first window from the list
    #     window_call = window_list[0]
    #     window_call.activate()
    #     pyautogui.hotkey('ctrl', 'shift', 'H')
    #     window_call.close()
    # elif exists_in_call_win :
    #     window_list = gw.getWindowsWithTitle(inCall_win_name)
    #     # Select the first window from the list
    #     window_call = window_list[0]
    #     window_call.activate()
    #     pyautogui.hotkey('ctrl', 'shift', 'H')
    #     window_call.close()
        
    # else:
    #     print('Check inCall window title ...\n')
    #     time.sleep(1);
        
   
     
    end_time = time.time()
    
    duration = int(end_time - start_time)
    
    print(f"Call and exectuion time around {duration} seconds...\n")
    

   
    return duration

####### ====================================================================
# Example usage for MPAT upload
def empty_folder(folder_path):
    # Check if the folder exists
    if not os.path.exists(folder_path):
        print(f"The folder '{folder_path}' does not exist.")
        return
    
    # Change permissions and empty the folder by deleting all its contents
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        os.chmod(item_path, stat.S_IWRITE)  # Change the permissions to writable
        if os.path.isdir(item_path):
            shutil.rmtree(item_path)
        else:
            os.remove(item_path)
    
    print(f"All contents of the folder '{folder_path}' have been deleted.")

    
def copy_folder_contents(source_folder, destination_folder):
    # Check if the source folder exists
    if not os.path.exists(source_folder):
        print(f"The source folder '{source_folder}' does not exist.")
        return
    
    # Create the destination folder if it does not exist
    os.makedirs(destination_folder, exist_ok=True)
    
    # Copy the contents of the source folder to the destination folder
    for item in os.listdir(source_folder):
        source_item = os.path.join(source_folder, item)
        destination_item = os.path.join(destination_folder, item)
        if os.path.isdir(source_item):
            shutil.copytree(source_item, destination_item)
        else:
            shutil.copy2(source_item, destination_item)
    
    print(f"All contents from '{source_folder}' have been copied to '{destination_folder}'.")

def change_permissions_and_delete(folder_path):
    # Change the permissions of the folder and its contents
    for root, dirs, files in os.walk(folder_path):
        for dir in dirs:
            os.chmod(os.path.join(root, dir), 0o777)
        for file in files:
            os.chmod(os.path.join(root, file), 0o777)
    # Delete the folder and its contents
    shutil.rmtree(folder_path)
    print(f"Deleted folder: {folder_path}")

def delete_subfolders_with_prefix(directory, prefix):
    # List all subfolders in the given directory
    for subfolder in os.listdir(directory):
        subfolder_path = os.path.join(directory, subfolder)
        # Check if the subfolder name starts with the given prefix and is a directory
        if os.path.isdir(subfolder_path) and subfolder.startswith(prefix):
            try:
                change_permissions_and_delete(subfolder_path)
            except PermissionError as e:
                print(f"PermissionError: {e}")
            except Exception as e:
                print(f"Error: {e}")

def activate_edge_and_search_keyword(winTitlewith='LogVisualizer',keyword="AAudio",keyword_2="OPENSLES"):
    # Get a list of all window titles
   
    
    while True:
        windows = gw.getAllTitles()

        # Filter windows that contain 'Microsoft Edge' in their title
        edge_windows = [title for title in windows if winTitlewith in title]
        

        if edge_windows:
            # Get the window object of the first matching window
            window = gw.getWindowsWithTitle(edge_windows[0])[0]
            print(f"Window ID: {window._hWnd}, Title: {window.title}")
        
            # Activate the window
            window.show()
            #window.maximize()
            window.activate()
        
            # Wait for the window to be activated
            time.sleep(2)
        
            # Send Ctrl+A to select all content
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
        
            # Send Ctrl+C to copy the selected content to the clipboard
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
        
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()
        
            # Count the occurrences of the keyword in the clipboard content
            keyword_count = clipboard_content.lower().count(keyword.lower())
            keyword_count_2 = clipboard_content.lower().count(keyword_2.lower())

            # Print the count of the keyword occurrences
            print(f"The keyword '{keyword}' appears {keyword_count} times in the page content.\n")
            print(f"The keyword '{keyword_2}' appears {keyword_count_2} times in the page content.\n")
            
            #window.close()
            break
        
        
        
    return window, keyword_count,keyword_count_2

def MTRw_logoff_cleanLog(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw):
    
    # Define the SSH connection details
    hostname = dut_ipaddr
    port = 22
    username = MTRw_usr
    password = MTRw_pw
    
    # Initialize the SSH client
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(hostname, port, username, password)
    
    # Define the PowerShell command
    #powershell_command = 'Get-Process'
    
    # Execute the PowerShell command
    stdin, stdout, stderr = client.exec_command(f'powershell.exe {"logoff console"}')
    
    time.sleep(3)
    
    stdin, stdout, stderr = client.exec_command(f'powershell.exe {"del c:/recordings/*.*"}')
    time.sleep(1)
    
    stdin, stdout, stderr = client.exec_command(f'powershell.exe {"del C:/Users/Skype/AppData/Local/Packages/MSTeamsRooms_8wekyb3d8bbwe/LocalCache/Microsoft/MSTeams/Logs/sc-tfw/mediastack/*.*"}')
    
    # Print the output
    print(stdout.read().decode())
    print(stderr.read().decode())
    
    # Close the connection
    client.close()


def auto_WinDesk_call_MTRw_TightVNC(call_Pos):
    
    start_time = time.time()
    
    #===============================================================================Prepare MTRw==============================
    MTRw_logoff_cleanLog(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw)
    
    while True:
        exists_TightVNC_win = window_exists_exact(TightVNC_win_name)
        
        if exists_TightVNC_win:
            window_list = gw.getWindowsWithTitle(TightVNC_win_name)
            TightVNC_window = window_list[0]
            time.sleep(3)
            #window_call.close()
            print(f"TightVNC Window with title '{TightVNC_win_name}' exists.")
            TightVNC_window.activate()
            TightVNC_window.maximize()
            break
        else:
            print(f"TightVNC Window with title '{TightVNC_win_name}' not exists.")
            time.sleep(2)
    if vqe_enable.upper() == "Y":
        MTRw_EnableVQE(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw)
    
    # MTRw Sign in
    new_mouse_x = TightVNC_window.left + int(call_Pos[1][1])
    new_mouse_y = TightVNC_window.top + int(call_Pos[1][2])
    
    pyautogui.moveTo(new_mouse_x, new_mouse_y)
    time.sleep(1)
    pyautogui.click()
    
    # wait enough time for MTRw to launch
    time.sleep(30)
    
          
    # Get all windows
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Print all window titles
    for item in window_list:
        if str.lower('Microsoft Teams')in item.title.lower():
            print(f"Teams Window found in {item.title}\n")
            if call_to_win_name.lower() in item.title.lower():
                item.close()
            elif inCall_win_name.lower() in item.title.lower():
                item.close()
            else:
                print('...')
                #item.activate()
                
                
            
    
    # Find the window by title
    window_list = gw.getWindowsWithTitle('Microsoft Teams')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams Main windows found\n")
        # Activate the window
        window.maximize()
    
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")
    
   
    #time.sleep(3)
    #window.activate()
    pyautogui.hotkey('ctrl', 'shift','N')
    
    time.sleep(5)

    time.sleep(1)
    
    window_list = gw.getWindowsWithTitle('New Message')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        
        print("Teams New Message windows found\n")
        # Activate the window
        window.maximize()
    
        window.activate()
    else:
        print("No window with the title 'Microsoft Teams' found.")

    pyautogui.typewrite(call_to_address)
    time.sleep(1)
 
    pyautogui.press('enter')
    # time.sleep(1)
    while True:
        exists_call_to_win = window_exists_exact(call_to_win_name)
        exists_in_call_win = window_exists_exact(inCall_win_name)
        
        if (exists_call_to_win or exists_in_call_win) :
            window_list = gw.getWindowsWithTitle(call_to_win_name)
            window_call = window_list[0]
            time.sleep(1)
            #window_call.close()
            print(f"Window with title '{call_to_win_name}' or '{inCall_win_name}' exists.")
            break
        else:
            print(f"Window with title '{call_to_win_name}' and '{inCall_win_name}' does not exist.")
            pyautogui.press('enter')
            time.sleep(2)
    
    pyautogui.press('enter')
    
    time.sleep(1)
    # Send Alt+Shift+V, start video call
    if call_audio_video.lower() in "video":
        pyautogui.hotkey('alt', 'shift', 'v')
    elif call_audio_video.lower() in "audio":
        pyautogui.hotkey('alt', 'shift', 'a')
    else:
        print('Please check input if "audio" or "video')
        
    
        
         
    
    time.sleep(3)
    
    #===============================================================================start Call==============================
  
    time.sleep(4)
    
    while True:
        exists_TightVNC_win = window_exists_exact(TightVNC_win_name)
        
        if exists_TightVNC_win:
            window_list = gw.getWindowsWithTitle(TightVNC_win_name)
            TightVNC_window = window_list[0]
            time.sleep(3)
            #window_call.close()
            print(f"TightVNC Window with title '{TightVNC_win_name}' exists.")
            TightVNC_window.activate()
            TightVNC_window.maximize()
            break
        else:
            print(f"TightVNC Window with title '{TightVNC_win_name}' not exists.")
            time.sleep(2)     
            
   
    time.sleep(1)
    # Calculate the new mouse position relative to the new window position
    new_mouse_x = TightVNC_window.left + int(call_Pos[2][1])
    new_mouse_y = TightVNC_window.top + int(call_Pos[2][2])
    
    pyautogui.moveTo(new_mouse_x, new_mouse_y)
    time.sleep(1)
    pyautogui.click()
    
    #===============================================================================End Call==============================
    time.sleep(call_Duration)  # call duration
    #===============================================================================End Call==============================
   
    time.sleep(2)
    # Calculate the new mouse position relative to the new window position
    new_mouse_x = TightVNC_window.left + int(call_Pos[3][1])
    new_mouse_y = TightVNC_window.top + int(call_Pos[3][2])
    
    pyautogui.moveTo(new_mouse_x, new_mouse_y)
    time.sleep(1)
    pyautogui.click()
    

    # Find the window by title
    window_list = gw.getWindowsWithTitle(call_to_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{call_to_win_name}" ...\n')
            print(item.title)
            if call_to_win_name.lower() in item.title.lower():
                item.activate()
                
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
            
    else:
        print(f"No window with the title '{call_to_win_name}' found...\n")
     
    time.sleep(1)
    
    window_list = gw.getWindowsWithTitle(inCall_win_name)
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        for item in window_list:
            print(f'try to End call on Windows Teams with "{inCall_win_name}" ...\n')
            print(item.title)
            if inCall_win_name.lower() in item.title.lower():
                item.activate()
                pyautogui.hotkey('ctrl', 'shift', 'H')
                pyautogui.hotkey('ctrl', 'shift', 'H')
                
                item.close()
    else:
        print(f"No window with the title '{inCall_win_name}' found...\n")
    
   # get MTRw logs
    transfer_logs_from_MTRw(transfer_logs_from_MTRw,local_path = local_path)
   
   
    end_time = time.time()
    
    duration = int(end_time - start_time)
    
    print(f"Call and exectuion time around {duration} seconds...\n")
    

   
    return duration

def Putty_run_alive(PuttyexePath = r"C:\Program Files\PuTTY"):
    
    currPath = os.getcwd()
    
    #TightVNCExePath = r"C:\TightVNC-win64-v2.1"
    os.chdir(PuttyexePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(2)
    
    cmdline1 = r'.\ssh_MTRw_NewTeams_sc-tfw_logoff.ps1'
     
    
    commands = cmdline1;
     
    #time.sleep(5)
    window_list = gw.getWindowsWithTitle('PowerShell')
    
    # Check if any windows were found
    if window_list:
        # Select the first window from the list
        window = window_list[0]
        window.activate()
        time.sleep(1)
        pyautogui.typewrite(commands)
        pyautogui.press('enter')

        #os.chdir(currPath);
        time.sleep(15)
        window.close()
    else:
        print("No window with the title 'PowerShell' found.")
        
    os.chdir(currPath)
    


def MPAT_MTRw_Test(ipaddr='10.172.206.65:5555', mpat_path=r'C:\Users\jianzou\OneDrive - Microsoft\MPAT_1.2.1',loop=1, testMethod = 'upload'.lower()):
    
    result_file = os.path.join(mpat_path,"result_url.txt")
    
    delete_subfolders_with_prefix(mpat_path,"result")
    # Empty the existing folder by deleting all its contents
   
# print all the test config and current time to txt
    # Get the current date and time
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if os.path.exists(result_file):
        # Append the clipboard content to the file
        with open(result_file, "a") as file:
            file.write(f"------------------------------------------{PATW_version}----------------------------------------------------\n")
            file.write(f"\nTest started on {current_datetime}\n\n")
            
            file.write(f"Call to address: {call_to_address}\n")
            
            file.write(f"MPAT Path: {MPAT_path}\n")
            
            #file.write(f"Putty exe Path: {Putty_path}\n")
            
            file.write(f"Audio or video call: {call_audio_video}\n")
            
            file.write(f"Call to Windows name keyword: {call_to_win_name}\n")
            
            file.write(f"In call Windows name keyword: {inCall_win_name}\n")
            
            file.write(f"DUT TightVNC Windows name: {TightVNC_win_name}\n")
             
            file.write(f"Call duration: {call_Duration}\n")
            
            file.write(f"DUT ip address: {dut_ipaddr}\n")
            
            file.write(f"VQE recording enabled?: {vqe_enable}\n")
            
            file.write(f"Test loops: {loopN}\n\n")

            #file.write(f"AAduio appears time: {times_AAudio}\n")
            #file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
            #file.write(lines[-3] + "\n")
           
    else:
        # Create a new file and write the clipboard content
        with open(result_file, "w") as file:
            file.write(f"\nTest started on {current_datetime}\n")
            
            file.write(f"Call to address: {call_to_address}\n")
            
            file.write(f"MPAT Path: {MPAT_path}\n")
            
            #file.write(f"Putty exe Path: {Putty_path}\n")
            
            
            file.write(f"Audio or video call: {call_audio_video}\n")
            
            file.write(f"Call to Windows name keyword: {call_to_win_name}\n")
            
            file.write(f"In call Windows name keyword: {inCall_win_name}\n")
            
            file.write(f"DUT TightVNC Windows name: {TightVNC_win_name}\n")
             
            file.write(f"Call duration: {call_Duration}\n")
            
            file.write(f"DUT ip address: {dut_ipaddr}\n")
            
            file.write(f"VQE recording enabled?: {vqe_enable}\n")
            
            file.write(f"Test loops: {loopN}\n\n")

            #file.write(f"AAduio appears time: {times_AAudio}\n")
   

    
    for ii in range(1,loopN + 1):
        loopStart_time = time.time()
        
        window_id = open_cmd_and_get_window_id(mpat_path)
        
        # existing_folder = os.path.join(mpat_path,"result")
        # if os.path.exists(existing_folder) and os.path.isdir(existing_folder):
        #     change_permissions_and_delete(existing_folder)
        
        
        if window_id:
            if testMethod.lower() == 'upload'.lower():
                Position = load_Call_Positions_from_csv(os.path.join(os.getcwd(),'MTRw_Positions.csv'))
                auto_WinDesk_call_MTRw_TightVNC(call_Pos = Position)
                
                # copy content from c:\MTRwLog to MPAT folder\result_ii
                copy_folder_contents(source_folder = local_path, destination_folder = os.path.join(MPAT_path, 'result_'+str(ii)))
                
                
            else:
                print("check testMethod input...\n")
            
            
            print(f"Command Prompt window ID: {window_id}\n")
            cmd_test = 'AudioToolClient.exe upload --path ' + local_path
            #window_id.maximize()
            window_id.activate()
            time.sleep(1)
            pyautogui.typewrite(cmd_test)
            time.sleep(1)
            
            pyautogui.press('enter')
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "The test results will be uploaded" not in clipboard_content:
                    time.sleep(1)
                else:
                    break
            pyautogui.typewrite('y')
            
            
            
            
           

            
            while True:
                # Send Ctrl +A copy all  the contents to clipboard
                window_id.activate()
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(0.2)
                # Get the content from the clipboard
                clipboard_content = pyperclip.paste()
                if "https://" not in clipboard_content:
                    time.sleep(3)
                else:
                    time.sleep(0.2)
                    break#time.sleep(10)
                
                    
            # while window_id.isActive:
            #     time.sleep(3)
            #     print('waiting...\n')
            
            #window_id.maximize()
            window_id.activate()
            
            time.sleep(1)

            
            # Send Ctrl +A copy all  the contents to clipboard
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()

            lines =  clipboard_content.splitlines(); 
            
            current_url = lines[-3]
            
            plot_data = extract_plots_from_mpat_url(url = current_url);

            #active Microsoft Edge Window and check AAudio and OpenSLES appears times
            #WinEdge,times_AAudio, times_OpenSLES = activate_edge_and_search_keyword(winTitlewith='LogVisualizer',keyword="AAudio",keyword_2="OPENSLES")
            time.sleep(3)  # wait enough seconds to get report full load to microsoft edge 
            #Edge_win_name='LogVisualizer';
            keyword="AAudio";
            keyword_2="OPENSLES"
  
            #windows = gw.getAllTitles()

            # Filter windows that contain 'Microsoft Edge' in their title
            #edge_windows = [title for title in windows if winTitlewith in title]
            
            while True:
                exists_IE_Win = window_exists_exact('LogVisualizer')
                
                
                if (exists_IE_Win) :
                    window_list = gw.getWindowsWithTitle('LogVisualizer')
                    window_IE = window_list[0]
                    time.sleep(1)
                    #window_call.close()
                    print(f"Window ID: {window_IE._hWnd}, Title: {window_IE.title}\n")
                    break
                else:
                    print(f"Window with title 'LogVisualizer' does not exist.")
                   
                    time.sleep(2)
                        
            
            print("MPAT result windows found\n")
            # Activate the window
            window_IE.show()
            #window.maximize()
            window_IE.activate()
        
            # Wait for the window to be activated
            time.sleep(1)
        
            # Send Ctrl+A to select all content
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.2)
        
            # Send Ctrl+C to copy the selected content to the clipboard
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.2)
        
            # Get the content from the clipboard
            clipboard_content = pyperclip.paste()
        
            # Count the occurrences of the keyword in the clipboard content
            times_AAudio = clipboard_content.lower().count(keyword.lower())
            times_OpenSLES = clipboard_content.lower().count(keyword_2.lower())
      

           
            #time.sleep(1)
                #window.close()
                

            # Save the clipboard content to a text file
            # Check if the file already exists
            if os.path.exists(result_file):
                # Append the clipboard content to the file
                with open(result_file, "a") as file:
                    file.write(f"Test {ii}\n")
                    #file.write(f"AAduio appears time: {times_AAudio}\n")
                    #file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
                    file.write(f"Len_Nearendinput0 = {plot_data[3]}, Len_Nearendoutput = {plot_data[4]}, Len_Farendinput = {plot_data[5]}\n")
                    file.write(f"Null_Nearendinput0 = {plot_data[6]}, Null_Nearendoutput = {plot_data[7]}, Null_Farendinput = {plot_data[8]}\n")
                    file.write(f"Avg_Nearendinput0 = {plot_data[9]:.1f}, Avg_Nearendoutput = {plot_data[10]:.1f}, Avg_Farendinput = {plot_data[11]:.1f}\n")
                    
                    
                    file.write(lines[-3] + "\n")
                   
            else:
                # Create a new file and write the clipboard content
                with open(result_file, "w") as file:
                    file.write(f"Test {ii}\n")
                    #file.write(f"AAduio appears time: {times_AAudio}\n")
                    #file.write(f"OpenSLES appears time: {times_OpenSLES}\n")
                    file.write(f"Len_Nearendinput0  = {plot_data[3]}, Len_Nearendoutput  = {plot_data[4]}, Len_Farendinput  = {plot_data[5]}\n")
                    file.write(f"Null_Nearendinput0 = {plot_data[6]}, Null_Nearendoutput = {plot_data[7]}, Null_Farendinput = {plot_data[8]}\n")
                    file.write(f"Avg_Nearendinput0  = {plot_data[9]:.1f}, Avg_Nearendoutput = {plot_data[10]:.1f}, Avg_Farendinput = {plot_data[11]:.1f}\n")
                    
                      
                    file.write(lines[-3] + "\n")
                    
            time.sleep(1)
            
            window_id.close()
        
                  
        else:
            print("Command Prompt window not found.")
        loop_time = time.time() - loopStart_time
        print(f"loop #{ii} test time {loop_time:.1f} seconds\n")
    
    # print end time
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if os.path.exists(result_file):
        # Append the clipboard content to the file
        with open(result_file, "a") as file:
            file.write(f"\nTest finished on {current_datetime}\n")
            

    else:
        # Create a new file and write the clipboard content
        with open(result_file, "w") as file:
            file.write(f"\nTest finished on {current_datetime}\n")
            
 
    return result_file

def open_notepad_with_file(file_name):
    # Check if the file exists
    if os.path.exists(file_name):
        # Open Notepad with the specified file
        subprocess.run(['notepad.exe', file_name])
    else:
        print(f"The file '{file_name}' does not exist.")
        
    return
def load_Call_Positions_from_csv(filename = r"MTRw_Positions.csv"):
    # Initialize an empty list to store the data
    data = []

    # Open the CSV file and read its content
    with open(filename, 'r', newline='') as csvfile:
        csvreader = csv.reader(csvfile)
    
        # Iterate over each row in the CSV file
        for row in csvreader:
            data.append(row)

    
    return data

if __name__ == "__main__":
    # print("test")
    MPAT_path = select_MPAT_folder()
    
    result_file = MPAT_MTRw_Test(ipaddr=dut_ipaddr,mpat_path= MPAT_path, loop=loopN, testMethod = 'upload')
    
    MTRw_DisableVQE(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw)
    
    open_notepad_with_file(result_file)
        
