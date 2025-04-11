# Print Job Monitoring System - Setup Guide

This document provides step-by-step instructions to set up and configure the Print Job Monitoring System on your servers.

## Prerequisites

Before you begin, ensure you have:

- A Windows server with PowerShell 5.1 or higher
- Administrative access to the print server(s) you want to monitor
- PowerShell Active Directory module installed
- Basic knowledge of PowerShell and Windows administration

## Installation Steps

### 1. Prepare the Files

1. Create a dedicated directory for the monitoring system, for example:
   ```
   C:\PrintMonitor
   ```

2. Save the two main files in this directory:
   - `MonitorPrintJobs.ps1` (PowerShell script)
   - `template.html` (HTML dashboard template)

### 2. Configure the Script

1. Open `MonitorPrintJobs.ps1` in a text editor (PowerShell ISE, VS Code, Notepad++, etc.)

2. Locate and modify the server configuration section at the beginning of the script:
   ```powershell
   $serverConfigs = @(
       @{
           ServerName = "your-print-server" # Change to your print server name
           Description = "Main Print Server" # Change description
       }
       # Add more server entries as needed
   )
   ```

3. For testing, you can use `localhost` if running directly on the print server

4. Optional: Customize other settings as needed:
   - `$queryInterval` - WMI query interval (in seconds)
   - `$htmlRefreshInterval` - How often to refresh the HTML report (in seconds)

### 3. Configure PowerShell Execution Policy

1. Open PowerShell as Administrator

2. Check your current execution policy:
   ```powershell
   Get-ExecutionPolicy
   ```

3. If needed, set a policy that allows running the script:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

### 4. Install Required Modules

1. Ensure the Active Directory module is installed:
   ```powershell
   Import-Module ActiveDirectory
   ```

2. If the module is not installed, you can add it on Windows Server:
   ```powershell
   Add-WindowsFeature RSAT-AD-PowerShell
   ```

### 5. Set Up WMI Access

1. Ensure WMI is properly configured on all print servers you will monitor

2. If monitoring remote servers, verify that WMI traffic is allowed through the firewall:
   ```powershell
   # Run on each print server
   netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes
   ```

### 6. Run the Script

1. Navigate to your installation directory:
   ```powershell
   cd C:\PrintMonitor
   ```

2. Run the script with administrative privileges:
   ```powershell
   .\MonitorPrintJobs.ps1
   ```

3. The script will start monitoring and provide output in the console:
   ```
   Monitoring Servers:
   - your-print-server (Main Print Server)
   Output Directory: C:\PrintMonitor
   HTML Template: C:\PrintMonitor\template.html
   HTML Report File: C:\PrintMonitor\PrintJobsLog-PrintServers-dd-MM-yyyy.html
   Initializing Runspace Pool...
   Starting initial monitoring threads...
   Attempting to start monitoring for server: your-print-server
   Successfully initiated monitoring for your-print-server.
   Monitoring started. Press Ctrl+C to stop.
   ```

### 7. Access the Dashboard

1. Open the generated HTML file in a web browser:
   ```
   C:\PrintMonitor\PrintJobsLog-PrintServers-[current-date].html
   ```

2. The dashboard will display:
   - Top users section
   - Top printers section
   - Detailed print jobs log

3. The dashboard will auto-refresh every 30 seconds by default

## Setting Up as a Scheduled Task

For continuous monitoring, you can set up the script as a scheduled task:

1. Open Task Scheduler in Windows

2. Create a new Task:
   - Name: Print Job Monitor
   - Run with highest privileges: Yes
   - Configure for: Your Windows version

3. Triggers:
   - Begin the task: At system startup
   - (Optional) Add additional triggers as needed

4. Actions:
   - Action: Start a program
   - Program/script: `powershell.exe`
   - Add arguments: `-ExecutionPolicy Bypass -File "C:\PrintMonitor\MonitorPrintJobs.ps1"`
   - Start in: `C:\PrintMonitor`

5. Conditions:
   - Start the task only if the computer is on AC power: Uncheck if on server
   
6. Settings:
   - Allow task to be run on demand: Yes
   - Run task as soon as possible after a scheduled start is missed: Yes
   - If the task fails, restart every: 1 minute
   - Attempt to restart up to: 3 times

## Setting Up as a Windows Service (Advanced)

For even more reliable operation, you can set up the script as a Windows Service using NSSM (Non-Sucking Service Manager):

1. Download NSSM from [nssm.cc](https://nssm.cc/)

2. Extract nssm.exe to a location on your server

3. Open Command Prompt as Administrator

4. Install the service:
   ```cmd
   nssm.exe install PrintJobMonitor
   ```

5. In the NSSM dialog:
   - Path: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
   - Startup directory: `C:\PrintMonitor`
   - Arguments: `-ExecutionPolicy Bypass -NoProfile -File "C:\PrintMonitor\MonitorPrintJobs.ps1"`
   - Service name: PrintJobMonitor
   - Set the service to Automatic start

6. Start the service:
   ```cmd
   nssm.exe start PrintJobMonitor
   ```

## GitHub Integration

To upload your project to GitHub:

1. Create a new repository on GitHub

2. Initialize a Git repository in your project folder:
   ```bash
   cd C:\PrintMonitor
   git init
   git add .
   git commit -m "Initial commit"
   ```

3. Connect to your GitHub repository:
   ```bash
   git remote add origin https://github.com/YourUsername/print-job-monitor.git
   git branch -M main
   git push -u origin main
   ```

## Customizing the Dashboard

You can customize the HTML template:

1. Open `template.html` in a text editor

2. Modify the header section to add your company name and logo:
   ```html
   <div class="logo-container">
       <img src="https://path.to/your-logo.png" alt="company logo" class="company-logo">
       <h1>Print Monitoring System</h1>
   </div>
   <h2>Your Company Name</h2>
   ```

3. Adjust CSS styles to match your company's color scheme in the `:root` variables

## Troubleshooting

### Script Fails to Start

1. Verify execution policy:
   ```powershell
   Get-ExecutionPolicy -List
   ```

2. Check for PowerShell module dependencies:
   ```powershell
   Import-Module ActiveDirectory -ErrorAction SilentlyContinue
   if ($?) { Write-Host "AD Module OK" } else { Write-Host "AD Module Missing" }
   ```

### No Data Appearing

1. Verify WMI is working:
   ```powershell
   Get-WmiObject -Class Win32_Printer -ComputerName your-print-server
   ```

2. Check print server accessibility:
   ```powershell
   Test-NetConnection -ComputerName your-print-server -Port 135
   ```

3. Verify print jobs are being processed:
   ```powershell
   Get-WmiObject -Class Win32_PrintJob -ComputerName your-print-server
   ```

### Additional Assistance

If you encounter issues not covered in this guide, please:

1. Check the console output for specific error messages
2. Review the script for any customizations needed for your environment
3. Reach out via the GitHub repository's Issues section
