Copilot generated ......

User Guide for auto_Teams_MTRw.py
Overview
The auto_Teams_MTRw.py script is designed to automate the process of making and managing calls using Microsoft Teams on a Windows PC. It includes functionalities for enabling and disabling VQE (Voice Quality Enhancement) recordings, transferring logs, and interacting with the Teams application through GUI automation.

Prerequisites
Before running the script, ensure the following:

Install MPAT on the x64 Windows PC.
Install Windows Teams and ensure it can call the DUT (Device Under Test) MTRw Teams.
Install TightVNC on both the Windows PC and MTRw, and ensure the Windows PC can connect to the DUT MTRw via TightVNC.
Ensure SSH connectivity from the Windows PC to the DUT MTRw Teams.
Script Configuration
The script prompts the user for various inputs to configure the test environment. Below are the prompts and their default values:

Test Method: Method for running the test (default: upload).
Call To Address: Email address to be called (default: sox_dut@MODERNCOMMS784471.onmicrosoft.com).
Call Type: Type of call (audio or video, default: video).
Call To Window Name: Keywords for the call window name (default: sox_dut).
In-Call Window Name: Keywords for the in-call window name (default: Microsoft Teams meeting).
TightVNC Window Title: Title of the TightVNC window (default: desktop-rfsncel).
Call Duration: Expected call duration in seconds (default: 10).
DUT IP Address: IP address of the DUT (default: 10.168.4.32).
MTRw Username: Administrator username for MTRw PC (default: admin).
MTRw Password: Password for MTRw PC (default: sfb).
Local Path: Local path to save MTRw logs (default: c:\MTRwLog).
Enable VQE Recording: Enable VQE Recording on MTRw (Y or N, default: Y).
Test Loop Count: Number of test loops (default: 3).
Main Functions
1. get_input(prompt, default_value)
Prompts the user for input and returns the entered value or the default value if no input is provided.

2. create_ssh_client(server, port, user, password)
Creates an SSH client connection to the specified server.

3. transfer_folder_from_remote(ssh_client, local_path, remote_path)
Transfers a folder from the remote server to the local machine.

4. transfer_logs_from_MTRw(remote_ip, local_path)
Transfers logs from the MTRw device to the local machine.

5. ensure_remote_folder_exists(ssh_client, remote_path)
Ensures that the specified remote folder exists, creating it if necessary.

6. transfer_folder_to_remote(ssh_client, local_path, remote_path)
Transfers a folder from the local machine to the remote server.

7. MTRw_EnableVQE(hostname, port, username, password)
Enables VQE recording on the MTRw device by transferring necessary files and executing PowerShell commands.

8. MTRw_DisableVQE(hostname, port, username, password)
Disables VQE recording on the MTRw device by executing PowerShell commands.

9. auto_WinDesk_call_MTRA()
Automates the process of making a call using Microsoft Teams on the Windows desktop.

10. auto_WinDesk_call_MTRw_TightVNC(call_Pos)
Automates the process of making a call using Microsoft Teams on the MTRw device via TightVNC.

11. MPAT_MTRw_Test(ipaddr, mpat_path, loop, testMethod)
Runs the MPAT test on the MTRw device, including making calls, transferring logs, and uploading results.

12. open_notepad_with_file(file_name)
Opens the specified file in Notepad.

13. load_Call_Positions_from_csv(filename)
Loads call positions from a CSV file.

Running the Script
Execute the Script: Run the script using Python:


Provide Inputs: Follow the prompts to provide the necessary inputs or accept the default values.

Monitor Execution: The script will automate the process of making calls, enabling/disabling VQE, transferring logs, and uploading results.

View Results: The results will be saved in the specified local path and can be viewed using Notepad.

Example Usage

This example demonstrates how to select the MPAT folder, run the MPAT test, disable VQE, and open the result file in Notepad.
