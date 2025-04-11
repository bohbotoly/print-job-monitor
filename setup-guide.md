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

1. Open `print-job-monitor.ps1` in a text editor (PowerShell ISE, VS Code, Notepad++, etc.)

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
   ```
   cd C:\PrintMonitor
   ```

2. Run the included batch file with administrative privileges (right-click and select "Run as administrator"):
   ```
   print-job-monitor.bat
   ```

   This batch file should be located in the same folder as your PowerShell script and HTML template.

3. The script will start monitoring and provide output in the console:
   ```
   ╔════════════════════════════════════════════════════════╗
   ║                                                        ║
   ║           Print Job Monitoring System v4.3             ║
   ║                   by bohbotoly                         ║
   ║                                                        ║
   ╚════════════════════════════════════════════════════════╝

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

   The batch file will automatically restart the monitoring script if it terminates for any reason.

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

## Setting Up as a Windows Service

For even more reliable operation, you can set up the script as a Windows Service using NSSM (Non-Sucking Service Manager):

1. Download NSSM from [nssm.cc](https://nssm.cc/)

2. Extract nssm.exe to a location on your server

3. Open Command Prompt as Administrator

4. Install the service:
   ```cmd
   nssm.exe install print-job-monitor
   ```

5. In the NSSM dialog:
   - Path: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
   - Startup directory: `C:\PrintMonitor`
   - Arguments: `-ExecutionPolicy Bypass -NoProfile -File "C:\PrintMonitor\print-job-monitor.ps1"`
   - Service name: PrintJobMonitor
   - Set the service to Automatic start

6. Start the service:
   ```cmd
   nssm.exe start PrintJobMonitor
   ```

## Advanced Customization Options

### Customizing Print Job Document Names

The script includes functionality to display more user-friendly document names for common print jobs. This is especially useful for system-generated documents that have cryptic names.

To customize document names:

1. Open the `template.html` file

2. Locate the `replaceDocumentNames()` JavaScript function:
   ```javascript
   function replaceDocumentNames() {
       const replacements = {
           "A2jn0293bdjOIDK": "Friendly Name 1",
           "A2jnjK2JbdjOIDK": "Friendly Name 2",
           "test page": "Friendly Name 3 - Test Page"
           // Add more known document names and their friendly versions
       };
       
       // Rest of the function...
   }
   ```

3. Add your own document name replacements to the `replacements` object:
   ```javascript
   const replacements = {
       "cryptic-document-name": "User-Friendly Document Name",
       "Excel.Sheet.12": "Excel Spreadsheet",
       "Microsoft Word Document": "Word Document",
       // Add more as needed
   };
   ```

4. The script will automatically replace these names in the dashboard display.

### Customizing Printer Name Formatting

The script includes special formatting for printer names:

1. In the PowerShell script, locate the `Build-HTMLContent` function

2. Find the printer name formatting section:
   ```powershell
   $displayPrinterName = if ($job.Printer -match '^(hp)') {
       $job.Printer.ToUpper()
   } else {
       $job.Printer
   }
   ```

3. You can modify this pattern to match your printer naming convention:
   ```powershell
   # Custom printer name formatting
   $displayPrinterName = if ($job.Printer -match '^(your-prefix)') {
       # Format printers with your prefix in a specific way
       $job.Printer.ToUpper()
   } elseif ($job.Printer -match 'another-pattern') {
       # Format printers matching another pattern differently
       "DEPT: " + $job.Printer
   } else {
       # Default formatting
       $job.Printer
   }
   ```

### Modifying AD User Information Display

You can customize what Active Directory information is displayed for users:

1. Locate the `Get-UserInfo` function in the PowerShell script:
   ```powershell
   function Get-UserInfo {
       param ([string]$SamAccountName)
       # Function code...
   }
   ```

2. Modify the `Get-ADUser` command to retrieve additional properties:
   ```powershell
   $user = Get-ADUser -Identity $normalizedKey -Properties DisplayName, Office, Department, Title -ErrorAction Stop
   ```

3. Update the returned user info object to include the new properties:
   ```powershell
   $userInfo = [PSCustomObject]@{
       DisplayName = $user.DisplayName
       Office = $user.Office
       Department = $user.Department
       Title = $user.Title
   }
   ```

4. Then update how this information is used in the `Build-HTMLContent` function:
   ```powershell
   $topUsersHtml += "<tr>
   <td class='highlight' title='$($userInfo.Office) - $($userInfo.Department)'>$crownIcon $($userInfo.DisplayName)</td>
   <td>$($userData.TotalJobs)</td>
   <td>$($userData.TotalPages)</td>
   </tr>"
   ```

### Adding Custom Metrics and Reports

You can add additional metrics to the dashboard:

1. Create new tracking collections at the beginning of the script:
   ```powershell
   $documentTypeCounts = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new()
   ```

2. Update the job processing in the runspace script block to track these metrics:
   ```powershell
   # Track document types
   $docType = if ($documentName -match '\.(\w+)$') { $matches[1].ToLower() } else { "unknown" }
   $null = $documentTypeCounts.AddOrUpdate(
       $docType,
       { [PSCustomObject]@{ Count = 1 } },
       { param($key, $existingValue) [PSCustomObject]@{ Count = $existingValue.Count + 1 } }
   )
   ```

3. Add new sections to the HTML template to display these metrics

4. Update the `Build-HTMLContent` function to populate these new sections

### Customizing the Dashboard Appearance

1. Open `template.html` in a text editor

2. To modify the color scheme, find the `:root` CSS variables section:
   ```css
   :root {
       --primary-color: #2763a2;
       --secondary-color: #1a1a2e;
       --accent-color: #ff9900;
       /* more colors... */
   }
   ```

3. Change these values to match your organization's branding

4. To add your company logo:
   ```html
   <div class="logo-container">
       <img src="path/to/your-logo.png" alt="company logo" class="company-logo">
       <h1>Print Monitoring System</h1>
   </div>
   <h2>Your Company Name</h2>
   ```

5. For a local image, place the logo file in the same directory and reference it:
   ```html
   <img src="company-logo.png" alt="company logo" class="company-logo">
   ```

### Modifying Refresh Intervals

1. To change how often the dashboard auto-refreshes:
   - In `template.html`, find the line `const refreshRate = 30000; // 30 seconds in milliseconds`
   - Change `30000` to your desired interval in milliseconds (e.g., `60000` for 1 minute)

2. To change how often the PowerShell script updates the HTML file:
   - In `MonitorPrintJobs.ps1`, find the line `$htmlRefreshInterval = 5  # Seconds between HTML refreshes`
   - Change `5` to your desired interval in seconds

### Adding Email Notifications

You can add email notification functionality for high-volume printing or other alerts:

1. Add the following variables to the configuration section:
   ```powershell
   $smtpServer = "your-smtp-server"
   $emailFrom = "printmonitor@yourdomain.com"
   $emailTo = "admin@yourdomain.com"
   $notificationThreshold = 100 # Pages threshold for notification
   ```

2. Create a function to send email notifications:
   ```powershell
   function Send-PrintNotification {
       param (
           [string]$UserName,
           [string]$PrinterName,
           [int]$Pages,
           [string]$DocumentName,
           [string]$ServerName
       )
       
       $subject = "High Volume Print Alert: $Pages pages"
       $body = @"
   High volume print job detected:
   
   User: $UserName
   Printer: $PrinterName
   Document: $DocumentName
   Pages: $Pages
   Server: $ServerName
   Time: $(Get-Date)
   "@
       
       try {
           Send-MailMessage -SmtpServer $smtpServer -From $emailFrom -To $emailTo -Subject $subject -Body $body
           Write-Host "Notification email sent for high volume print job." -ForegroundColor Yellow
       } catch {
           Write-Warning "Failed to send notification email: $($_.Exception.Message)"
       }
   }
   ```

3. Add the notification trigger in the print job processing section:
   ```powershell
   if ($pageCount -ge $notificationThreshold) {
       Send-PrintNotification -UserName $userName -PrinterName $printerName -Pages $pageCount -DocumentName $documentName -ServerName $ServerName
   }
   ```

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

### Frequent Script Crashes

1. Verify server stability and resource availability
2. Check Windows Event Logs for related errors
3. Increase error handling robustness:
   ```powershell
   # Add to the configuration section
   $maxRetries = 5
   $retryDelay = 10 # seconds
   
   # Then modify WMI connection code to use retry logic
   $retryCount = 0
   $connected = $false
   while (-not $connected -and $retryCount -lt $maxRetries) {
       try {
           $scope.Connect()
           $connected = $true
       } catch {
           $retryCount++
           Write-Warning "Failed to connect to WMI on attempt $retryCount. Retrying in $retryDelay seconds..."
           Start-Sleep -Seconds $retryDelay
       }
   }
   ```

### Additional Assistance

If you encounter issues not covered in this guide, please:

1. Check the console output for specific error messages
2. Review the script for any customizations needed for your environment
3. Reach out via the GitHub repository's Issues section
