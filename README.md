# Print Job Monitoring System

A PowerShell-based solution for monitoring print jobs across one or multiple print servers with a dynamic HTML dashboard.

![Print Job Monitor Dashboard](https://i.postimg.cc/VLgHmH19/print-monitor-dashboard.png)

## Features

- Real-time monitoring of print jobs across multiple print servers
- Responsive HTML dashboard with auto-refresh capabilities
- Track top users and printers by print volume
- Comprehensive job history with filtering and sorting
- Export data to Excel with one click
- User-friendly interface with dark mode design
- Resilient multi-threaded architecture with automatic recovery

## Requirements

- Windows Server with Print Services role
- PowerShell 5.1 or higher
- Active Directory module for PowerShell
- Administrative access to print servers

## Files

- `MonitorPrintJobs.bat` - Bat to Run the Script
- `MonitorPrintJobs.ps1` - Main PowerShell script
- `template.html` - HTML dashboard template

## Installation

1. Clone or download this repository to your print server or monitoring server
2. Ensure both files are in the same directory
3. Modify the server configuration in the script (see Configuration section)
4. Run the script with administrative privileges

## Configuration

Edit the `$serverConfigs` section at the beginning of the PowerShell script to define which print servers to monitor:

```powershell
$serverConfigs = @(
    @{
        ServerName = "PrintServer1" # Your print server name
        Description = "Main Print Server" # Description
    },
    @{
        ServerName = "PrintServer2" # Add more servers as needed
        Description = "Secondary Print Server"
    }
)
```

For local testing, you can use `"localhost"` as the ServerName.

## Usage

### Starting the Monitor

Run the script with administrative privileges:

```powershell
.\MonitorPrintJobs.ps1
```

The script will:
1. Start monitoring print jobs on the configured servers
2. Generate an HTML report in the same directory
3. Update the report at regular intervals

### Viewing the Dashboard

Open the generated HTML file (named `PrintJobsLog-PrintServers-[date].html`) in any modern web browser.

The dashboard features:
- Auto-refresh every 30 seconds (can be toggled on/off)
- Tables of top users and printers
- Complete print job history
- Search and filtering options
- Export to Excel functionality

## Dashboard Controls

- **Auto Refresh** - Toggles automatic page refresh every 30 seconds
- **Manually Refresh** - Forces an immediate refresh of the data
- **Export to Excel** - Exports all print job data to an Excel file
- **Show Full Table** - Toggles between showing all jobs or just the most recent 300

## Customization

### HTML Template

You can customize the HTML template (`template.html`) to:
- Change the company logo and name
- Adjust colors and styling
- Modify the dashboard layout
- Add additional information sections

### Script Parameters

Key parameters you might want to adjust:
- `$queryInterval` - How frequently to poll for new print jobs (in seconds)
- `$htmlRefreshInterval` - How often to update the HTML report (in seconds)

## Troubleshooting

### Common Issues

1. **Script fails to connect to print server**
   - Ensure you have administrative access to the print server
   - Verify that WMI is accessible on the target server
   - Check firewall settings

2. **No print jobs appear in the report**
   - Verify that print services are running on the server
   - Confirm that print jobs are being processed
   - Check the server name in the configuration

3. **Script crashes or stops**
   - The script has built-in recovery for most errors
   - Check the PowerShell console for error messages
   - Increase the logging level for more detailed troubleshooting

### Logging

The script outputs status messages to the console with color coding:
- Green: Successful operations
- Yellow: Warnings or state changes
- Red: Errors
- Cyan: Informational messages

## Advanced Configuration

### Active Directory Integration

The script uses Active Directory to resolve user information. If you're using this in a non-AD environment, you may need to modify the `Get-UserInfo` function.

### Multiple Print Servers

For monitoring many print servers, consider adjusting the `$maxThreads` parameter to optimize performance:

```powershell
$maxThreads = [number of servers] + 1
```

## License

[Your License Here]

## Author

[Your Name/Organization]

