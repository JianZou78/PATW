import pygetwindow as gw
import pyautogui
import keyboard
import time
import csv
import os
import subprocess
import tkinter as tk
from tkinter import filedialog
import paramiko
from scp import SCPClient

def get_input(prompt, default_value):
    user_input = input(f"{prompt} (default: {default_value}): ")
    return user_input if user_input else default_value

dut_ipaddr = get_input('Please input the ip address of the DUT, for example: 10.168.4.32\n','10.168.4.32')

MTRw_usr = get_input('Please input the administrator to login MTRw PC, for example: admin\n','admin')

MTRw_pw = get_input('Please input the password to login MTRw PC, for example: sfb\n','sfb')

def capture_mouse_position_in_tightVNC(Win_Title='desktop-rfsncel'):
    # Get the tightVNC window position and size
    tightVNC_window = gw.getWindowsWithTitle(Win_Title)[0]
    
    tightVNC_window.maximize()
    window_x, window_y = tightVNC_window.left, tightVNC_window.top
    window_width, window_height = tightVNC_window.width, tightVNC_window.height

    # Get the mouse position relative to the screen
    mouse_x, mouse_y = pyautogui.position()

    # Calculate the mouse position relative to the tightVNC window
    relative_x = mouse_x - window_x
    relative_y = mouse_y - window_y

    if 0 <= relative_x <= window_width and 0 <= relative_y <= window_height:
        return relative_x, relative_y
    else:
        return None

def get_tightVNC_window_and_mouse(x, y,Win_Title='desktop-rfsncel'):
    # Get the tightVNC window
    tightVNC_window = gw.getWindowsWithTitle(Win_Title)[0]

    # Calculate the new mouse position relative to the new window position
    new_mouse_x = tightVNC_window.left + x
    new_mouse_y = tightVNC_window.top + y

    # Move the mouse to the new position
    pyautogui.moveTo(new_mouse_x, new_mouse_y)
    
def select_Putty_folder():
    
    print (" select the Putty folder path....\n" )
    root = tk.Tk()
    
    root.withdraw()
     
    # Open a folder dialog and get the selected folder path
    
    Putty_path = filedialog.askdirectory()
     
    # Print the selected folder path
    
    print(f"Selected folder: {Putty_path}")
    return Putty_path

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


def Putty_run_alive(PuttyexePath = r"C:\Program Files\PuTTY"):
    
    currPath = os.getcwd()
    
    #TightVNCExePath = r"C:\TightVNC-win64-v2.1"
    os.chdir(PuttyexePath)
    
    subprocess.run(["powershell", "-Command", "Start-Process powershell"])

    time.sleep(2)
    
    cmdline1 = r'.\ssh_MTRw_NewTeams_sc-tfw_logoff.ps1'
     
    
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
    
    tightVNC_win_name = get_input('Please input tightVNC window title, for example: desktop-rfsncel\n','desktop-rfsncel')

    #Putty_path = select_Putty_folder()
    #Putty_run_alive(Putty_path)
    MTRw_logoff_cleanLog(hostname=dut_ipaddr,port=22,username = MTRw_usr, password = MTRw_pw)

    print("Move mouse to the tightVNC window Skype Sign in icon, Press CTRL+A to capture the mouse position inside the given window...")


    keyboard.wait('ctrl+a')

# Capture the mouse position inside the tightVNC window to sign in
    captured_position_signin = capture_mouse_position_in_tightVNC(Win_Title=tightVNC_win_name)
    print(f"Mouse position captured at: ({captured_position_signin[0]}, {captured_position_signin[1]}) inside the given window")

# wait enough time for MTRw to launch
    time.sleep(15)   
    
    print("Make a call to MTRw, Move mouse to the tightVNC window Accept call icon, Press CTRL+A to capture the mouse position inside the given window...")
    
    keyboard.wait('ctrl+a')
# Capture the mouse position inside the tightVNC window to accept call
    captured_position_call = capture_mouse_position_in_tightVNC(Win_Title=tightVNC_win_name)
    
    print(f"Mouse position captured at: ({captured_position_call[0]}, {captured_position_call[1]}) inside the given window")

# Capture the mouse position inside the tightVNC window to end the call

    print("Move mouse to the EndCall icon, Press CTRL+A to capture the mouse position inside the given window...\n")

    keyboard.wait('ctrl+a')

    captured_position_Endcall = capture_mouse_position_in_tightVNC(Win_Title=tightVNC_win_name)
    print(f"Mouse position captured at: ({captured_position_Endcall[0]}, {captured_position_Endcall[1]}) inside the given window\n")

    Postion_file = os.path.join(os.getcwd(),'MTRw_Positions.csv')
# Check if the file exists
    file_exists = os.path.isfile(Postion_file)
    

    with open(Postion_file, 'w', newline='') as csvfile:
        
        csvwriter = csv.writer(csvfile)
        
        csvwriter.writerow(['Item','X_relative','Y_relative'])
        csvwriter.writerow(['Sign in',captured_position_signin[0],captured_position_signin[1]])
        csvwriter.writerow(['Call',captured_position_call[0],captured_position_call[1]])
        csvwriter.writerow(['EndCall',captured_position_Endcall[0],captured_position_Endcall[1]])

