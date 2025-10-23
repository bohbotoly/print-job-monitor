<#
.SYNOPSIS
    Monitors print jobs across Windows print servers and generates real-time HTML reports.

.DESCRIPTION
    The Print Job Monitor provides real-time monitoring of print jobs across one or more Windows print servers.
    It uses WMI event watchers to detect new print jobs as they occur, tracks user and printer statistics,
    and generates a continuously updated HTML report with the following information:

    - Top users by pages printed
    - Top printers by usage
    - Recent print job history with user, document, and printer details

    The script uses Active Directory integration to enrich data with user display names and office locations,
    and employs parallel processing via PowerShell runspaces for efficient multi-server monitoring.

.PARAMETER ServerConfigs
    Array of hashtables containing server configurations. Each hashtable should have:
    - ServerName: The name or IP of the print server to monitor
    - Description: A friendly description of the server

.PARAMETER HtmlTemplatePath
    Path to the HTML template file. If not specified, defaults to 'template.html' in the script directory.

.PARAMETER OutputDirectory
    Directory where HTML reports will be saved. If not specified, defaults to the script directory.

.PARAMETER WmiQueryInterval
    WMI query polling interval in seconds. Lower values increase responsiveness but also CPU load.
    Default: 1 second

.PARAMETER HtmlRefreshInterval
    Seconds between HTML report refreshes. Default: 5 seconds

.PARAMETER RunspaceCheckInterval
    Seconds between checking runspace health status. Default: 2 seconds

.PARAMETER MaxConcurrentThreads
    Maximum number of concurrent monitoring threads. Default: 2

.PARAMETER TopItemsCount
    Number of top users and printers to display in the report. Default: 10

.PARAMETER MaxDocumentNameLength
    Maximum length for document names in the report before truncation. Default: 40 characters

.PARAMETER FileRetryCount
    Number of retry attempts when writing to HTML file. Default: 3

.PARAMETER FileRetryDelayMs
    Delay in milliseconds between file write retry attempts. Default: 300ms

.PARAMETER LogLevel
    Logging verbosity level: None, Error, Warning, Information, Verbose. Default: Information

.EXAMPLE
    .\print-job-monitor.ps1
    Monitors localhost using default settings.

.EXAMPLE
    .\print-job-monitor.ps1 -ServerConfigs @(@{ServerName='PRINT01';Description='Main Office'}) -Verbose
    Monitors PRINT01 server with verbose output enabled.

.EXAMPLE
    $servers = @(
        @{ServerName='PRINT01'; Description='Building A'},
        @{ServerName='PRINT02'; Description='Building B'}
    )
    .\print-job-monitor.ps1 -ServerConfigs $servers -LogLevel Verbose

.NOTES
    File Name      : print-job-monitor.ps1
    Author         : Print Monitoring Team
    Prerequisite   : PowerShell 5.1 or higher, Active Directory module
    Copyright      : (c) 2025. All rights reserved.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/
#>

[CmdletBinding(DefaultParameterSetName='Default')]
param(
    [Parameter(Mandatory=$false, HelpMessage='Array of server configuration hashtables')]
    [ValidateNotNullOrEmpty()]
    [hashtable[]]$ServerConfigs = @(
        @{
            ServerName = 'localhost'
            Description = 'Local Print Server'
        }
    ),

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if (-not $_ -or (Test-Path -Path (Split-Path -Parent $_) -PathType Container)) {
            $true
        } else {
            throw "Parent directory of HtmlTemplatePath does not exist: $_"
        }
    })]
    [string]$HtmlTemplatePath,

    [Parameter(Mandatory=$false)]
    [ValidateScript({
        if (-not $_ -or (Test-Path -Path $_ -PathType Container)) {
            $true
        } else {
            throw "OutputDirectory does not exist: $_"
        }
    })]
    [string]$OutputDirectory,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 60)]
    [int]$WmiQueryInterval = 1,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 300)]
    [int]$HtmlRefreshInterval = 5,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 60)]
    [int]$RunspaceCheckInterval = 2,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 10)]
    [int]$MaxConcurrentThreads = 2,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 100)]
    [int]$TopItemsCount = 10,

    [Parameter(Mandatory=$false)]
    [ValidateRange(10, 200)]
    [int]$MaxDocumentNameLength = 40,

    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 10)]
    [int]$FileRetryCount = 3,

    [Parameter(Mandatory=$false)]
    [ValidateRange(100, 5000)]
    [int]$FileRetryDelayMs = 300,

    [Parameter(Mandatory=$false)]
    [ValidateSet('None', 'Error', 'Warning', 'Information', 'Verbose')]
    [string]$LogLevel = 'Information'
)

#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Import Active Directory module if available
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Warning "Active Directory module not available. User lookups will use usernames only."
    Write-Warning "Error: $($_.Exception.Message)"
}

#region Script Configuration

# Script-scoped configuration object
$script:Config = [PSCustomObject]@{
    # Paths
    ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
    HtmlTemplatePath = $null
    OutputDirectory = $null
    HtmlFile = $null

    # Server Configuration
    ServerConfigs = $ServerConfigs

    # Timing Configuration
    WmiQueryInterval = $WmiQueryInterval
    HtmlRefreshInterval = $HtmlRefreshInterval
    RunspaceCheckInterval = $RunspaceCheckInterval

    # Display Configuration
    TopItemsCount = $TopItemsCount
    MaxDocumentNameLength = $MaxDocumentNameLength
    TopUserIcon = '👑'  # Crown emoji (U+1F451)
    TopPrinterIcon = '👑'  # Crown emoji (U+1F451)

    # File Operation Configuration
    FileRetryCount = $FileRetryCount
    FileRetryDelayMs = $FileRetryDelayMs

    # Printer Name Patterns
    PrinterNameUppercasePatterns = @('^HP', '^bspr')

    # Localization
    NoJobsMessage = 'לא נמצאו הדפסות ביום הנוכחי'

    # Threading
    MaxConcurrentThreads = $MaxConcurrentThreads

    # Logging
    LogLevel = $LogLevel

    # Runtime State
    CurrentDate = Get-Date -Format 'yyyy-MM-dd'
    LastHtmlRefresh = [datetime]::MinValue
    LastRunspaceCheck = [datetime]::MinValue
}

# Set paths with fallbacks
$script:Config.OutputDirectory = if ($OutputDirectory) {
    $OutputDirectory
} else {
    $script:Config.ScriptDirectory
}

$script:Config.HtmlTemplatePath = if ($HtmlTemplatePath) {
    $HtmlTemplatePath
} else {
    Join-Path -Path $script:Config.ScriptDirectory -ChildPath 'template.html'
}

# Ensure output directory exists
if (-not (Test-Path -Path $script:Config.OutputDirectory)) {
    New-Item -ItemType Directory -Path $script:Config.OutputDirectory -Force | Out-Null
}

# Generate HTML output filename base (will be made unique in Start-PrintJobMonitoring)
$dateString = Get-Date -Format 'dd-MM-yyyy'
$script:HtmlFileBase = Join-Path -Path $script:Config.OutputDirectory -ChildPath "PrintJobsLog-PrintServers-$dateString"

# Thread-safe collections
$script:AdCache = @{}
$script:PrinterCache = @{}
$script:RecentJobs = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$script:UserPrintCounts = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:PrinterPrintCounts = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:SyncHash = [hashtable]::Synchronized(@{})
$script:SyncHash.MessageQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()

# Active runspace tracking
$script:ActiveRunspaces = @{}
$script:RunspacePool = $null

#endregion

#region Logging Functions

<#
.SYNOPSIS
    Writes a log message with specified severity level.

.DESCRIPTION
    Provides structured logging with configurable severity levels. Messages are written to
    appropriate PowerShell streams based on severity.

.PARAMETER Message
    The message to log.

.PARAMETER Level
    The severity level: Verbose, Information, Warning, or Error.

.PARAMETER ErrorRecord
    Optional ErrorRecord object for error-level logging.

.EXAMPLE
    Write-Log -Message "Processing started" -Level Information

.EXAMPLE
    Write-Log -Message "Failed to connect" -Level Error -ErrorRecord $_
#>
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet('Verbose', 'Information', 'Warning', 'Error')]
        [string]$Level = 'Information',

        [Parameter(Mandatory=$false)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    # Check if we should output based on configured log level
    $logLevels = @{
        'None' = 0
        'Error' = 1
        'Warning' = 2
        'Information' = 3
        'Verbose' = 4
    }

    $currentLevel = $logLevels[$script:Config.LogLevel]
    $messageLevel = $logLevels[$Level]

    if ($messageLevel -gt $currentLevel) {
        return
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formattedMessage = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'Verbose' {
            Write-Verbose -Message $formattedMessage
        }
        'Information' {
            Write-Information -MessageData $formattedMessage -InformationAction Continue
        }
        'Warning' {
            Write-Warning -Message $formattedMessage
        }
        'Error' {
            if ($ErrorRecord) {
                # When ErrorRecord is provided, just output the formatted message as the error text
                # Don't use -ErrorRecord parameter to avoid parameter set conflicts
                Write-Error -Message "$formattedMessage`nDetails: $($ErrorRecord.Exception.Message)" -Category $ErrorRecord.CategoryInfo.Category
            } else {
                Write-Error -Message $formattedMessage
            }
        }
    }
}

#endregion

#region Utility Functions

<#
.SYNOPSIS
    Generates a unique filename by appending a counter if file exists.

.DESCRIPTION
    Checks if a file exists at the specified path and appends a numeric counter
    to create a unique filename if necessary.

.PARAMETER BaseName
    The base name of the file without extension.

.PARAMETER Extension
    The file extension without the dot.

.OUTPUTS
    System.String. The unique filename with full path.

.EXAMPLE
    Get-UniqueFileName -BaseName "C:\Logs\report" -Extension "html"
    Returns "C:\Logs\report.html" or "C:\Logs\report.01.html" if file exists.
#>
function Get-UniqueFileName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Extension
    )

    $counter = 0
    $newFileName = "$BaseName.$Extension"

    while (Test-Path -Path $newFileName -PathType Leaf) {
        $counter++
        $newFileName = "$BaseName.$($counter.ToString('D2')).$Extension"
    }

    Write-Log -Message "Generated unique filename: $newFileName" -Level Verbose
    return $newFileName
}

<#
.SYNOPSIS
    Retrieves and caches Active Directory user information.

.DESCRIPTION
    Looks up user information from Active Directory and caches the results to minimize
    repeated AD queries. Returns display name and office location.

.PARAMETER SamAccountName
    The SAM account name of the user to look up.

.OUTPUTS
    PSCustomObject with DisplayName and Office properties.

.EXAMPLE
    $userInfo = Get-UserInfo -SamAccountName "jsmith"
#>
function Get-UserInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$SamAccountName
    )

    # Normalize the key (remove domain prefix, convert to lowercase)
    $normalizedKey = ($SamAccountName -split '\\' | Select-Object -Last 1).ToLowerInvariant()

    # Return cached value if available
    if ($script:AdCache.ContainsKey($normalizedKey)) {
        Write-Log -Message "Retrieved cached AD info for user: $normalizedKey" -Level Verbose
        return $script:AdCache[$normalizedKey]
    }

    try {
        Write-Log -Message "Querying Active Directory for user: $normalizedKey" -Level Verbose
        $user = Get-ADUser -Identity $normalizedKey -Properties DisplayName, Office -ErrorAction Stop

        $userInfo = [PSCustomObject]@{
            DisplayName = $user.DisplayName
            Office = $user.Office
        }

        Write-Log -Message "Successfully retrieved AD info for user: $normalizedKey" -Level Verbose
    }
    catch {
        Write-Log -Message "Failed to retrieve AD info for user: $normalizedKey. Error: $($_.Exception.Message)" -Level Warning

        # Cache failed lookups to avoid repeated attempts
        $userInfo = [PSCustomObject]@{
            DisplayName = $SamAccountName
            Office = 'Unknown'
        }
    }

    # Add to cache
    $script:AdCache[$normalizedKey] = $userInfo
    return $userInfo
}

<#
.SYNOPSIS
    Retrieves and caches printer information.

.DESCRIPTION
    Queries the print server for printer details and caches the results to minimize
    repeated queries. Returns printer name and related information.

.PARAMETER PrinterName
    The name of the printer to look up.

.PARAMETER ServerName
    The print server hosting the printer.

.OUTPUTS
    PSCustomObject with PrinterName property.

.EXAMPLE
    $printerInfo = Get-PrinterInfo -PrinterName "HP-LaserJet" -ServerName "PRINT01"
#>
function Get-PrinterInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$PrinterName,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$ServerName
    )

    # Create cache key: servername:::printername (lowercase)
    $cacheKey = "$ServerName:::$PrinterName".ToLowerInvariant()

    # Return cached value if available
    if ($script:PrinterCache.ContainsKey($cacheKey)) {
        Write-Log -Message "Retrieved cached printer info: $cacheKey" -Level Verbose
        return $script:PrinterCache[$cacheKey]
    }

    try {
        Write-Log -Message "Querying printer details: $PrinterName on $ServerName" -Level Verbose
        $printer = Get-Printer -Name $PrinterName -ComputerName $ServerName -ErrorAction Stop

        $printerInfo = [PSCustomObject]@{
            PrinterName = $printer.Name
        }

        $script:PrinterCache[$cacheKey] = $printerInfo
        Write-Log -Message "Successfully retrieved printer info: $cacheKey" -Level Verbose
        return $printerInfo
    }
    catch {
        Write-Log -Message "Failed to retrieve printer '$PrinterName' on server '$ServerName'. Error: $($_.Exception.Message)" -Level Warning

        # Cache failed lookups
        $printerInfo = [PSCustomObject]@{
            PrinterName = $PrinterName
        }

        $script:PrinterCache[$cacheKey] = $printerInfo
        return $printerInfo
    }
}

<#
.SYNOPSIS
    Writes content to a file with retry logic.

.DESCRIPTION
    Attempts to write content to a file with configurable retry logic to handle
    transient file system issues. Uses UTF-8 encoding with BOM for proper character support.

.PARAMETER Path
    The full path to the file to write.

.PARAMETER Content
    The content to write to the file.

.PARAMETER RetryCount
    Number of retry attempts. Default from configuration.

.PARAMETER DelayMilliseconds
    Delay between retry attempts in milliseconds. Default from configuration.

.OUTPUTS
    System.Boolean. True if write succeeded, False otherwise.

.EXAMPLE
    $success = Write-FileWithRetry -Path "C:\Reports\output.html" -Content $htmlContent
#>
function Write-FileWithRetry {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Content,

        [Parameter(Mandatory=$false)]
        [ValidateRange(1, 10)]
        [int]$RetryCount = $script:Config.FileRetryCount,

        [Parameter(Mandatory=$false)]
        [ValidateRange(100, 10000)]
        [int]$DelayMilliseconds = $script:Config.FileRetryDelayMs
    )

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            # Ensure parent directory exists
            $parentDir = Split-Path -Path $Path -Parent
            if (-not (Test-Path -Path $parentDir -PathType Container)) {
                New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
                Write-Log -Message "Created directory: $parentDir" -Level Verbose
            }

            # Write with UTF-8 encoding (with BOM for Hebrew text support)
            $utf8WithBom = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllText($Path, $Content, $utf8WithBom)

            Write-Log -Message "Successfully wrote file: $Path" -Level Verbose
            return $true
        }
        catch {
            $message = "Attempt $attempt of $RetryCount failed to write file '$Path'. Error: $($_.Exception.Message)"

            if ($attempt -lt $RetryCount) {
                Write-Log -Message $message -Level Warning
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
            else {
                Write-Log -Message "Failed to write file '$Path' after $RetryCount attempts." -Level Error -ErrorRecord $_
            }
        }
    }

    return $false
}

<#
.SYNOPSIS
    Applies printer name formatting rules.

.DESCRIPTION
    Converts printer names to uppercase if they match configured patterns.

.PARAMETER PrinterName
    The printer name to format.

.OUTPUTS
    System.String. The formatted printer name.

.EXAMPLE
    $formatted = Format-PrinterName -PrinterName "hp-laserjet-01"
    Returns "HP-LASERJET-01"
#>
function Format-PrinterName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$PrinterName
    )

    foreach ($pattern in $script:Config.PrinterNameUppercasePatterns) {
        if ($PrinterName -match $pattern) {
            return $PrinterName.ToUpper()
        }
    }

    return $PrinterName
}

#endregion

#region HTML Generation

<#
.SYNOPSIS
    Builds the complete HTML report content.

.DESCRIPTION
    Generates the HTML report by processing user statistics, printer statistics,
    and recent print jobs. Replaces placeholders in the HTML template with actual data.

.PARAMETER UserPrintCounts
    Concurrent dictionary containing user print statistics.

.PARAMETER PrinterPrintCounts
    Concurrent dictionary containing printer print statistics.

.PARAMETER PrintJobs
    Concurrent queue containing recent print job records.

.PARAMETER TemplatePath
    Path to the HTML template file.

.OUTPUTS
    System.String. The complete HTML content ready to write to file.

.EXAMPLE
    $html = Build-HtmlContent -UserPrintCounts $userCounts -PrinterPrintCounts $printerCounts -PrintJobs $jobs -TemplatePath $templatePath
#>
function Build-HtmlContent {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$PrintJobs,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$TemplatePath
    )

    Write-Log -Message "Building HTML content from template: $TemplatePath" -Level Verbose

    # Read HTML template
    try {
        $htmlTemplate = Get-Content -Path $TemplatePath -Raw -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Log -Message "Failed to read HTML template file: $($_.Exception.Message)" -Level Error -ErrorRecord $_

        # Fallback minimal template
        $htmlTemplate = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <title>Print Job Monitor - Error</title>
</head>
<body>
    <h1>Error: Could not load template</h1>
    <p>Print monitoring data is still being collected.</p>
</body>
</html>
"@
    }

    # Build top users HTML
    $topUsersHtml = Build-TopUsersHtml -UserPrintCounts $UserPrintCounts

    # Build top printers HTML
    $topPrintersHtml = Build-TopPrintersHtml -PrinterPrintCounts $PrinterPrintCounts

    # Build print jobs HTML
    $printJobsHtml = Build-PrintJobsHtml -PrintJobs $PrintJobs

    # Replace placeholders
    $htmlContent = $htmlTemplate -replace '\{\{topUsersHtml\}\}', $topUsersHtml
    $htmlContent = $htmlContent -replace '\{\{topPrintersHtml\}\}', $topPrintersHtml
    $htmlContent = $htmlContent -replace '\{\{printJobsHtml\}\}', $printJobsHtml

    Write-Log -Message "HTML content built successfully" -Level Verbose
    return $htmlContent
}

<#
.SYNOPSIS
    Builds the HTML table rows for top users.

.DESCRIPTION
    Creates HTML table rows showing top users by pages printed, including crown icon for top user.

.PARAMETER UserPrintCounts
    Concurrent dictionary containing user print statistics.

.OUTPUTS
    System.String. HTML table rows for top users.
#>
function Build-TopUsersHtml {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts
    )

    $htmlBuilder = [System.Text.StringBuilder]::new()

    # Get all entries and filter/sort safely
    $topUsers = @($UserPrintCounts.GetEnumerator() |
        Where-Object { $_.Value -ne $null } |
        Sort-Object {
            try {
                # Use simple property access with default value
                if ($null -ne $_.Value.TotalPages) {
                    [int]$_.Value.TotalPages
                } else {
                    0
                }
            } catch {
                0
            }
        } -Descending |
        Select-Object -First $script:Config.TopItemsCount)

    if ($topUsers.Count -eq 0) {
        return $htmlBuilder.ToString()
    }

    $topUserKey = $topUsers[0].Key

    foreach ($userEntry in $topUsers) {
        $userKey = $userEntry.Key
        $userData = $userEntry.Value

        # Defensive check - skip if data is incomplete
        if (-not $userData) {
            continue
        }

        # Try to access properties safely
        try {
            $totalJobs = $userData.TotalJobs
            $totalPages = $userData.TotalPages

            # Skip if values are null or zero (invalid data)
            if ($null -eq $totalJobs -or $null -eq $totalPages) {
                continue
            }

            $crownIcon = if ($userKey -eq $topUserKey) { $script:Config.TopUserIcon } else { '' }

            $userInfo = Get-UserInfo -SamAccountName $userKey

            $null = $htmlBuilder.AppendLine("<tr>")
            $null = $htmlBuilder.AppendLine("    <td class='highlight' title='$($userInfo.Office)'>$crownIcon $($userInfo.DisplayName)</td>")
            $null = $htmlBuilder.AppendLine("    <td>$totalJobs</td>")
            $null = $htmlBuilder.AppendLine("    <td>$totalPages</td>")
            $null = $htmlBuilder.AppendLine("</tr>")
        }
        catch {
            # Skip entries with property access errors
            continue
        }
    }

    return $htmlBuilder.ToString()
}

<#
.SYNOPSIS
    Builds the HTML table rows for top printers.

.DESCRIPTION
    Creates HTML table rows showing top printers by pages printed, including crown icon for top printer.

.PARAMETER PrinterPrintCounts
    Concurrent dictionary containing printer print statistics.

.OUTPUTS
    System.String. HTML table rows for top printers.
#>
function Build-TopPrintersHtml {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts
    )

    $htmlBuilder = [System.Text.StringBuilder]::new()

    # Get all entries and filter/sort safely
    $topPrinters = @($PrinterPrintCounts.GetEnumerator() |
        Where-Object { $_.Value -ne $null } |
        Sort-Object {
            try {
                # Use simple property access with default value
                if ($null -ne $_.Value.TotalPages) {
                    [int]$_.Value.TotalPages
                } else {
                    0
                }
            } catch {
                0
            }
        } -Descending |
        Select-Object -First $script:Config.TopItemsCount)

    if ($topPrinters.Count -eq 0) {
        return $htmlBuilder.ToString()
    }

    $topPrinterKey = $topPrinters[0].Key

    foreach ($printerEntry in $topPrinters) {
        $printerKey = $printerEntry.Key
        $printerData = $printerEntry.Value

        # Defensive check - skip if data is incomplete
        if (-not $printerData) {
            continue
        }

        # Try to access properties safely
        try {
            $totalJobs = $printerData.TotalJobs
            $totalPages = $printerData.TotalPages

            # Skip if values are null or zero (invalid data)
            if ($null -eq $totalJobs -or $null -eq $totalPages) {
                continue
            }

            $crownIcon = if ($printerKey -eq $topPrinterKey) { $script:Config.TopPrinterIcon } else { '' }

            # Extract server and printer name from combined key (server:::printer)
            $parts = $printerKey -split ':::'
            $printerNameOnly = if ($parts.Count -ge 2) { $parts[1] } else { $printerKey }
            $serverNameOnly = if ($parts.Count -ge 2) { $parts[0] } else { 'Unknown' }

            $displayPrinterName = Format-PrinterName -PrinterName $printerNameOnly

            $null = $htmlBuilder.AppendLine("<tr>")
            $null = $htmlBuilder.AppendLine("    <td class='highlight'>$displayPrinterName $crownIcon</td>")
            $null = $htmlBuilder.AppendLine("    <td>$totalJobs</td>")
            $null = $htmlBuilder.AppendLine("    <td>$totalPages</td>")
            $null = $htmlBuilder.AppendLine("    <td>$serverNameOnly</td>")
            $null = $htmlBuilder.AppendLine("</tr>")
        }
        catch {
            # Skip entries with property access errors
            continue
        }
    }

    return $htmlBuilder.ToString()
}

<#
.SYNOPSIS
    Builds the HTML table rows for recent print jobs.

.DESCRIPTION
    Creates HTML table rows showing recent print jobs with user, document, printer, and timestamp information.

.PARAMETER PrintJobs
    Concurrent queue containing recent print job records.

.OUTPUTS
    System.String. HTML table rows for print jobs.
#>
function Build-PrintJobsHtml {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$PrintJobs
    )

    $htmlBuilder = [System.Text.StringBuilder]::new()

    # Convert queue to array and sort by time (newest first)
    $jobArray = $PrintJobs.ToArray() | Sort-Object {
        if ($_.Time -match '(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})') {
            Get-Date -Year $matches[3] -Month $matches[2] -Day $matches[1] -Hour $matches[4] -Minute $matches[5]
        }
        else {
            [DateTime]::MinValue
        }
    } -Descending

    # Track unique jobs to prevent duplicates
    $processedJobs = New-Object System.Collections.Generic.HashSet[string]

    foreach ($job in $jobArray) {
        # Create unique job identifier
        $uniqueJobId = if ($job.JobKey) {
            $job.JobKey
        } else {
            "$($job.User)-$($job.JobId)-$($job.Time)"
        }

        # Skip if already processed
        if (-not $processedJobs.Add($uniqueJobId)) {
            continue
        }

        # Get user information
        $userInfo = Get-UserInfo -SamAccountName $job.User

        # Truncate long document names
        $documentName = $job.Document
        $documentTooltip = ''

        if ($documentName.Length -gt $script:Config.MaxDocumentNameLength) {
            $documentName = $documentName.Substring(0, $script:Config.MaxDocumentNameLength - 3) + '...'
            $documentTooltip = "title='$($job.Document)'"
        }

        # Format printer name
        $displayPrinterName = Format-PrinterName -PrinterName $job.Printer

        $null = $htmlBuilder.AppendLine("<tr>")
        $null = $htmlBuilder.AppendLine("    <td>$($job.Time)</td>")
        $null = $htmlBuilder.AppendLine("    <td title='$($userInfo.Office)'>$($userInfo.DisplayName)</td>")
        $null = $htmlBuilder.AppendLine("    <td $documentTooltip>$documentName</td>")
        $null = $htmlBuilder.AppendLine("    <td>$($job.Pages)</td>")
        $null = $htmlBuilder.AppendLine("    <td>$displayPrinterName</td>")
        $null = $htmlBuilder.AppendLine("    <td>$($job.Server)</td>")
        $null = $htmlBuilder.AppendLine("</tr>")
    }

    # Add fallback message if no jobs found
    if ($htmlBuilder.Length -eq 0) {
        $null = $htmlBuilder.AppendLine("<tr>")
        $null = $htmlBuilder.AppendLine("    <td colspan='6' class='highlight' style='text-align:center;'>$($script:Config.NoJobsMessage)</td>")
        $null = $htmlBuilder.AppendLine("</tr>")
    }

    return $htmlBuilder.ToString()
}

<#
.SYNOPSIS
    Updates the HTML report file if the refresh interval has elapsed.

.DESCRIPTION
    Checks if enough time has passed since the last HTML update and regenerates
    the HTML file if necessary.

.PARAMETER LastRefreshTime
    DateTime of the last HTML refresh.

.PARAMETER RefreshIntervalSeconds
    Minimum seconds between refreshes.

.PARAMETER OutputFile
    Path to the HTML output file.

.PARAMETER UserCounts
    User print statistics dictionary.

.PARAMETER PrinterCounts
    Printer print statistics dictionary.

.PARAMETER JobsQueue
    Recent print jobs queue.

.PARAMETER HtmlTemplatePath
    Path to HTML template file.

.OUTPUTS
    System.DateTime. The updated last refresh time.

.EXAMPLE
    $lastRefresh = Update-HtmlIfNeeded -LastRefreshTime $lastRefresh -RefreshIntervalSeconds 5 -OutputFile $htmlFile -UserCounts $userCounts -PrinterCounts $printerCounts -JobsQueue $recentJobs -HtmlTemplatePath $templatePath
#>
function Update-HtmlIfNeeded {
    [CmdletBinding()]
    [OutputType([datetime])]
    param(
        [Parameter(Mandatory=$true)]
        [datetime]$LastRefreshTime,

        [Parameter(Mandatory=$true)]
        [ValidateRange(1, 300)]
        [int]$RefreshIntervalSeconds,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputFile,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$JobsQueue,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$HtmlTemplatePath
    )

    $currentTime = Get-Date
    $elapsedSeconds = ($currentTime - $LastRefreshTime).TotalSeconds

    if ($elapsedSeconds -ge $RefreshIntervalSeconds) {
        Write-Log -Message "HTML refresh interval reached ($elapsedSeconds seconds). Generating HTML report." -Level Verbose

        try {
            $htmlContent = Build-HtmlContent `
                -UserPrintCounts $UserCounts `
                -PrinterPrintCounts $PrinterCounts `
                -PrintJobs $JobsQueue `
                -TemplatePath $HtmlTemplatePath

            if (Write-FileWithRetry -Path $OutputFile -Content $htmlContent) {
                Write-Log -Message "HTML report updated successfully: $OutputFile" -Level Information
                return $currentTime
            }
            else {
                Write-Log -Message "Failed to write HTML report after retries: $OutputFile" -Level Warning
                return $LastRefreshTime
            }
        }
        catch {
            Write-Log -Message "Error generating HTML report: $($_.Exception.Message)" -Level Error -ErrorRecord $_
            return $LastRefreshTime
        }
    }

    return $LastRefreshTime
}

#endregion

#region Parallel Monitoring

<#
.SYNOPSIS
    Script block that runs in a separate runspace to monitor a print server.

.DESCRIPTION
    This script block is executed in parallel for each print server being monitored.
    It creates a WMI event watcher for print job events and processes them in real-time.
#>
$script:MonitoringScriptBlock = {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServerName,

        [Parameter(Mandatory=$true)]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,

        [Parameter(Mandatory=$true)]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,

        [Parameter(Mandatory=$true)]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$RecentJobs,

        [Parameter(Mandatory=$true)]
        [int]$WmiQueryInterval,

        [Parameter(Mandatory=$true)]
        [hashtable]$SyncHash
    )

    # Function to send messages back to main thread
    function Send-MessageToMainThread {
        param(
            [string]$Message,
            [string]$Color = 'White'
        )

        $SyncHash.MessageQueue.Enqueue(@{
            Message = $Message
            Color = $Color
            Timestamp = Get-Date
        })
    }

    # Initialize stop flag
    $stopRequested = $false
    $SyncHash["StopFlag-$ServerName"] = [ref]$stopRequested

    # Register stop event handler
    $eventAction = {
        param($Sender, $EventArgs)

        $serverName = $EventArgs.MessageData.ServerName
        $stopFlag = $EventArgs.MessageData.StopFlag
        $sendMessageFunction = $EventArgs.MessageData.SendMessageFunction

        & $sendMessageFunction -Message "[$serverName] Stop event received." -Color 'Yellow'
        $stopFlag.Value = $true
    }

    $messageData = @{
        ServerName = $ServerName
        StopFlag = [ref]$stopRequested
        SendMessageFunction = ${function:Send-MessageToMainThread}
    }

    $null = Register-EngineEvent -SourceIdentifier "StopRunspace-$ServerName" -Action $eventAction -MessageData $messageData

    # Initialize WMI watcher
    $watcher = $null
    $options = New-Object System.Management.EventWatcherOptions
    $options.Timeout = [TimeSpan]::FromSeconds(2)

    try {
        Send-MessageToMainThread -Message "[$ServerName] Initializing WMI event watcher..." -Color 'Cyan'

        $query = "SELECT * FROM __InstanceCreationEvent WITHIN $WmiQueryInterval WHERE TargetInstance ISA 'Win32_PrintJob'"
        $scope = New-Object System.Management.ManagementScope("\\$ServerName\root\cimv2")
        $scope.Connect()

        $watcher = New-Object System.Management.ManagementEventWatcher($scope, $query)
        $watcher.Options = $options

        Send-MessageToMainThread -Message "[$ServerName] WMI watcher initialized successfully. Starting monitoring loop..." -Color 'Green'

        # Main event processing loop
        while (-not $stopRequested) {
            try {
                if ($stopRequested) {
                    Send-MessageToMainThread -Message "[$ServerName] Stop requested before WaitForNextEvent." -Color 'Yellow'
                    break
                }

                $event = $watcher.WaitForNextEvent()

                if ($stopRequested) {
                    Send-MessageToMainThread -Message "[$ServerName] Stop requested after WaitForNextEvent." -Color 'Yellow'
                    if ($event) { $event.Dispose() }
                    break
                }

                if ($event) {
                    $job = $event.TargetInstance

                    # Extract job properties with defensive checks
                    $printerName = if ($job.Name) {
                        ($job.Name -split ',' | Select-Object -First 1)
                    } else {
                        'UnknownPrinter'
                    }

                    $userName = if ($job.Owner) { $job.Owner } else { 'UnknownUser' }
                    $documentName = if ($job.Document) { $job.Document } else { 'UnknownDocument' }
                    $timeStamp = Get-Date -Format 'dd-MM-yyyy HH:mm'
                    $pageCount = if ($job.TotalPages -and $job.TotalPages -gt 0) { $job.TotalPages } else { 1 }
                    $jobId = if ($job.JobId) { $job.JobId } else { [Guid]::NewGuid().ToString() }

                    Send-MessageToMainThread -Message "[$ServerName] New Print Job: User=$userName, Printer=$printerName, Pages=$pageCount, Document=$documentName" -Color 'Green'

                    # Normalize keys for dictionary operations
                    $normalizedUserName = ($userName -split '\\' | Select-Object -Last 1).ToLowerInvariant()
                    $printerKey = "$ServerName:::$printerName".ToLowerInvariant()
                    $uniqueId = [Guid]::NewGuid().ToString()
                    $jobKey = "$normalizedUserName-$jobId-$uniqueId"

                    # Update user statistics
                    $null = $UserPrintCounts.AddOrUpdate(
                        $normalizedUserName,
                        {
                            [PSCustomObject]@{
                                TotalJobs = 1
                                TotalPages = $pageCount
                            }
                        },
                        {
                            param($key, $existingValue)
                            [PSCustomObject]@{
                                TotalJobs = $existingValue.TotalJobs + 1
                                TotalPages = $existingValue.TotalPages + $pageCount
                            }
                        }
                    )

                    # Update printer statistics
                    $null = $PrinterPrintCounts.AddOrUpdate(
                        $printerKey,
                        {
                            [PSCustomObject]@{
                                TotalJobs = 1
                                TotalPages = $pageCount
                            }
                        },
                        {
                            param($key, $existingValue)
                            [PSCustomObject]@{
                                TotalJobs = $existingValue.TotalJobs + 1
                                TotalPages = $existingValue.TotalPages + $pageCount
                            }
                        }
                    )

                    # Add to recent jobs queue
                    $newJobEntry = @{
                        Server = $ServerName
                        Printer = $printerName
                        Pages = $pageCount
                        Document = $documentName
                        User = $userName
                        Time = $timeStamp
                        JobId = $jobId
                        JobKey = $jobKey
                    }

                    $RecentJobs.Enqueue($newJobEntry)
                    Send-MessageToMainThread -Message "[$ServerName] Job processed successfully: JobID=$jobId" -Color 'Cyan'

                    $event.Dispose()
                }
            }
            catch {
                # Expected timeout is normal behavior
                if ($_.Exception.Message -like '*Timed out*') {
                    continue
                }
                else {
                    Send-MessageToMainThread -Message "[$ServerName] Error in event processing loop: $($_.Exception.Message)" -Color 'Red'

                    if ($stopRequested) {
                        break
                    }

                    Start-Sleep -Seconds 2
                }
            }
        }
    }
    catch {
        Send-MessageToMainThread -Message "[$ServerName] Failed to initialize WMI watcher: $($_.Exception.Message)" -Color 'Red'
        throw
    }
    finally {
        Send-MessageToMainThread -Message "[$ServerName] Runspace shutting down..." -Color 'Yellow'

        if ($watcher) {
            Send-MessageToMainThread -Message "[$ServerName] Stopping and disposing WMI watcher..." -Color 'Yellow'
            try { $watcher.Stop() } catch { }
            try { $watcher.Dispose() } catch { }
        }

        Unregister-EngineEvent -SourceIdentifier "StopRunspace-$ServerName" -ErrorAction SilentlyContinue
        Send-MessageToMainThread -Message "[$ServerName] Runspace cleanup completed." -Color 'Yellow'
    }
}

<#
.SYNOPSIS
    Starts a monitoring runspace for a single print server.

.DESCRIPTION
    Creates and initializes a PowerShell runspace that monitors print jobs on the specified server.

.PARAMETER ServerConfig
    Hashtable containing ServerName and Description.

.PARAMETER RunspacePool
    The runspace pool to use for the new runspace.

.PARAMETER UserPrintCounts
    Shared user statistics dictionary.

.PARAMETER PrinterPrintCounts
    Shared printer statistics dictionary.

.PARAMETER RecentJobs
    Shared recent jobs queue.

.PARAMETER WmiQueryInterval
    WMI query polling interval in seconds.

.PARAMETER SyncHash
    Synchronized hashtable for cross-thread communication.

.OUTPUTS
    PSCustomObject containing runspace information and status.

.EXAMPLE
    $runspaceInfo = Start-ServerMonitorRunspace -ServerConfig $config -RunspacePool $pool -UserPrintCounts $userCounts -PrinterPrintCounts $printerCounts -RecentJobs $jobs -WmiQueryInterval 1 -SyncHash $syncHash
#>
function Start-ServerMonitorRunspace {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [hashtable]$ServerConfig,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$RecentJobs,

        [Parameter(Mandatory=$true)]
        [ValidateRange(1, 60)]
        [int]$WmiQueryInterval,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [hashtable]$SyncHash
    )

    $serverName = $ServerConfig.ServerName
    Write-Log -Message "Starting monitoring runspace for server: $serverName" -Level Information

    try {
        $powershell = [powershell]::Create()
        $null = $powershell.AddScript($script:MonitoringScriptBlock).AddParameters(@{
            ServerName = $serverName
            UserPrintCounts = $UserPrintCounts
            PrinterPrintCounts = $PrinterPrintCounts
            RecentJobs = $RecentJobs
            WmiQueryInterval = $WmiQueryInterval
            SyncHash = $SyncHash
        })

        $powershell.RunspacePool = $RunspacePool
        $handle = $powershell.BeginInvoke()

        Write-Log -Message "Successfully started monitoring for server: $serverName" -Level Information

        return [PSCustomObject]@{
            PowerShell = $powershell
            Handle = $handle
            ServerName = $serverName
            StartTime = Get-Date
            Status = 'Running'
            LastError = $null
        }
    }
    catch {
        Write-Log -Message "Failed to start runspace for server '$serverName': $($_.Exception.Message)" -Level Error -ErrorRecord $_

        return [PSCustomObject]@{
            PowerShell = $null
            Handle = $null
            ServerName = $serverName
            StartTime = Get-Date
            Status = 'FailedToStart'
            LastError = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Monitors runspace health and restarts failed runspaces.

.DESCRIPTION
    Checks the status of all active runspaces and automatically restarts any that have failed or completed unexpectedly.

.PARAMETER ActiveRunspaces
    Hashtable of active runspace information objects.

.PARAMETER ServerConfigs
    Array of server configuration hashtables.

.PARAMETER RunspacePool
    The runspace pool.

.PARAMETER UserPrintCounts
    Shared user statistics dictionary.

.PARAMETER PrinterPrintCounts
    Shared printer statistics dictionary.

.PARAMETER RecentJobs
    Shared recent jobs queue.

.PARAMETER WmiQueryInterval
    WMI query polling interval.

.PARAMETER SyncHash
    Synchronized hashtable for communication.

.EXAMPLE
    Test-RunspaceHealth -ActiveRunspaces $activeRunspaces -ServerConfigs $serverConfigs -RunspacePool $pool -UserPrintCounts $userCounts -PrinterPrintCounts $printerCounts -RecentJobs $jobs -WmiQueryInterval 1 -SyncHash $syncHash
#>
function Test-RunspaceHealth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [hashtable]$ActiveRunspaces,

        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [hashtable[]]$ServerConfigs,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Management.Automation.Runspaces.RunspacePool]$RunspacePool,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$RecentJobs,

        [Parameter(Mandatory=$true)]
        [ValidateRange(1, 60)]
        [int]$WmiQueryInterval,

        [Parameter(Mandatory=$true)]
        [ValidateNotNull()]
        [hashtable]$SyncHash
    )

    Write-Log -Message 'Checking runspace health status...' -Level Verbose

    foreach ($serverName in @($ActiveRunspaces.Keys)) {
        $runspaceInfo = $ActiveRunspaces[$serverName]

        if ($runspaceInfo.Status -ne 'Running') {
            continue
        }

        if ($runspaceInfo.Handle -and $runspaceInfo.Handle.IsCompleted) {
            Write-Log -Message "Runspace for server '$serverName' completed unexpectedly." -Level Warning

            try {
                $null = $runspaceInfo.PowerShell.EndInvoke($runspaceInfo.Handle)
                $runspaceInfo.Status = 'CompletedUnexpectedly'
                Write-Log -Message "Runspace for '$serverName' ended without error (unexpected). Restarting..." -Level Warning
            }
            catch {
                $runspaceInfo.Status = 'Failed'
                $runspaceInfo.LastError = $_.Exception.Message
                Write-Log -Message "Runspace for server '$serverName' failed: $($_.Exception.Message)" -Level Error -ErrorRecord $_
            }
            finally {
                try { $runspaceInfo.PowerShell.Dispose() } catch { }
                $runspaceInfo.PowerShell = $null
                $runspaceInfo.Handle = $null
            }

            # Attempt restart
            Write-Log -Message "Attempting to restart monitoring for server: $serverName" -Level Information

            $serverConfig = $ServerConfigs | Where-Object { $_.ServerName -eq $serverName } | Select-Object -First 1

            if ($serverConfig) {
                $newRunspaceInfo = Start-ServerMonitorRunspace `
                    -ServerConfig $serverConfig `
                    -RunspacePool $RunspacePool `
                    -UserPrintCounts $UserPrintCounts `
                    -PrinterPrintCounts $PrinterPrintCounts `
                    -RecentJobs $RecentJobs `
                    -WmiQueryInterval $WmiQueryInterval `
                    -SyncHash $SyncHash

                $ActiveRunspaces[$serverName] = $newRunspaceInfo
            }
            else {
                Write-Log -Message "Could not find configuration for server '$serverName' to restart monitoring." -Level Error
            }
        }
    }
}

#endregion

#region Main Execution

<#
.SYNOPSIS
    Main execution function that orchestrates the print job monitoring.

.DESCRIPTION
    Initializes the monitoring infrastructure, starts runspaces for each server,
    and manages the main monitoring loop with periodic HTML updates and runspace health checks.
#>
function Start-PrintJobMonitoring {
    [CmdletBinding()]
    param()

    # Validate prerequisites
    Write-Log -Message 'Validating prerequisites...' -Level Information

    if (-not (Test-Path -Path $script:Config.HtmlTemplatePath -PathType Leaf)) {
        Write-Log -Message "HTML template file not found: $($script:Config.HtmlTemplatePath)" -Level Error
        throw "Required HTML template file not found: $($script:Config.HtmlTemplatePath)"
    }

    # Generate unique HTML filename
    $script:Config.HtmlFile = Get-UniqueFileName -BaseName $script:HtmlFileBase -Extension 'html'

    # Display configuration
    Write-Log -Message '=== Print Job Monitor Starting ===' -Level Information
    Write-Log -Message "Monitoring Servers:" -Level Information
    foreach ($config in $script:Config.ServerConfigs) {
        Write-Log -Message "  - $($config.ServerName) ($($config.Description))" -Level Information
    }
    Write-Log -Message "Output Directory: $($script:Config.OutputDirectory)" -Level Information
    Write-Log -Message "HTML Template: $($script:Config.HtmlTemplatePath)" -Level Information
    Write-Log -Message "HTML Report File: $($script:Config.HtmlFile)" -Level Information
    Write-Log -Message "Log Level: $($script:Config.LogLevel)" -Level Information

    try {
        # Initialize runspace pool
        Write-Log -Message "Initializing runspace pool with $($script:Config.MaxConcurrentThreads) threads..." -Level Information

        $script:RunspacePool = [runspacefactory]::CreateRunspacePool(1, $script:Config.MaxConcurrentThreads)
        $script:RunspacePool.Open()

        # Start monitoring for each server
        Write-Log -Message 'Starting monitoring threads for configured servers...' -Level Information

        foreach ($serverConfig in $script:Config.ServerConfigs) {
            $runspaceInfo = Start-ServerMonitorRunspace `
                -ServerConfig $serverConfig `
                -RunspacePool $script:RunspacePool `
                -UserPrintCounts $script:UserPrintCounts `
                -PrinterPrintCounts $script:PrinterPrintCounts `
                -RecentJobs $script:RecentJobs `
                -WmiQueryInterval $script:Config.WmiQueryInterval `
                -SyncHash $script:SyncHash

            $script:ActiveRunspaces[$serverConfig.ServerName] = $runspaceInfo
        }

        Write-Log -Message 'All monitoring threads started successfully. Press Ctrl+C to stop.' -Level Information

        # Main monitoring loop
        while ($true) {
            try {
                # Process messages from background threads
                while ($script:SyncHash.MessageQueue.Count -gt 0) {
                    $message = $null
                    if ($script:SyncHash.MessageQueue.TryDequeue([ref]$message)) {
                        Write-Log -Message $message.Message -Level Information
                    }
                }

                # Check for date change (log rotation)
                $today = Get-Date -Format 'yyyy-MM-dd'
                if ($today -ne $script:Config.CurrentDate) {
                    Write-Log -Message "Date changed to $today. Stopping monitoring for log rotation." -Level Warning
                    break
                }

                # Periodic runspace health check
                if ((Get-Date) -ge $script:Config.LastRunspaceCheck.AddSeconds($script:Config.RunspaceCheckInterval)) {
                    Test-RunspaceHealth `
                        -ActiveRunspaces $script:ActiveRunspaces `
                        -ServerConfigs $script:Config.ServerConfigs `
                        -RunspacePool $script:RunspacePool `
                        -UserPrintCounts $script:UserPrintCounts `
                        -PrinterPrintCounts $script:PrinterPrintCounts `
                        -RecentJobs $script:RecentJobs `
                        -WmiQueryInterval $script:Config.WmiQueryInterval `
                        -SyncHash $script:SyncHash

                    $script:Config.LastRunspaceCheck = Get-Date
                }

                # Periodic HTML update
                $script:Config.LastHtmlRefresh = Update-HtmlIfNeeded `
                    -LastRefreshTime $script:Config.LastHtmlRefresh `
                    -RefreshIntervalSeconds $script:Config.HtmlRefreshInterval `
                    -OutputFile $script:Config.HtmlFile `
                    -UserCounts $script:UserPrintCounts `
                    -PrinterCounts $script:PrinterPrintCounts `
                    -JobsQueue $script:RecentJobs `
                    -HtmlTemplatePath $script:Config.HtmlTemplatePath

                # Sleep to prevent high CPU usage
                Start-Sleep -Milliseconds 100
            }
            catch {
                # Log errors in the monitoring loop but don't exit
                Write-Log -Message "Error in monitoring loop: $($_.Exception.Message)" -Level Error
                Write-Log -Message "Stack trace: $($_.ScriptStackTrace)" -Level Verbose

                # Sleep before continuing to prevent rapid error loops
                Start-Sleep -Seconds 1
            }
        }
    }
    finally {
        # Cleanup
        Write-Log -Message 'Stopping monitoring and cleaning up resources...' -Level Information

        # Stop all active runspaces
        foreach ($serverName in @($script:ActiveRunspaces.Keys)) {
            $runspaceInfo = $script:ActiveRunspaces[$serverName]

            if ($runspaceInfo.PowerShell -and $runspaceInfo.Status -eq 'Running') {
                Write-Log -Message "Stopping runspace for server: $serverName" -Level Information

                try {
                    # Signal stop via event
                    New-Event -SourceIdentifier "StopRunspace-$serverName" -EventArguments @() | Out-Null

                    # Set stop flag directly
                    $stopFlagRef = $script:SyncHash["StopFlag-$serverName"]
                    if ($stopFlagRef) {
                        $stopFlagRef.Value = $true
                        Write-Log -Message "Stop flag set for server: $serverName" -Level Verbose
                    }
                }
                catch {
                    Write-Log -Message "Error triggering stop event for '$serverName': $($_.Exception.Message)" -Level Warning
                }

                # Allow graceful shutdown
                Start-Sleep -Seconds 2

                # Force stop if still running
                try {
                    if (-not $runspaceInfo.Handle.IsCompleted) {
                        Write-Log -Message "Forcefully stopping runspace for server: $serverName" -Level Warning
                        $runspaceInfo.PowerShell.Stop()
                    }
                    else {
                        Write-Log -Message "Runspace for '$serverName' stopped gracefully." -Level Information
                    }
                }
                catch {
                    Write-Log -Message "Error stopping PowerShell for '$serverName': $($_.Exception.Message)" -Level Warning
                }
            }

            # Dispose resources
            try {
                if ($runspaceInfo.PowerShell) {
                    if ($runspaceInfo.Handle -and -not $runspaceInfo.Handle.IsCompleted) {
                        try { $null = $runspaceInfo.PowerShell.EndInvoke($runspaceInfo.Handle) } catch { }
                    }
                    $runspaceInfo.PowerShell.Dispose()
                }
            }
            catch {
                Write-Log -Message "Error disposing PowerShell for '$serverName': $($_.Exception.Message)" -Level Warning
            }
        }

        # Clean up stop flags
        foreach ($key in @($script:SyncHash.Keys)) {
            if ($key.StartsWith('StopFlag-')) {
                try {
                    $script:SyncHash.Remove($key)
                }
                catch {
                    Write-Log -Message "Error removing stop flag '$key': $($_.Exception.Message)" -Level Warning
                }
            }
        }

        # Close runspace pool
        if ($script:RunspacePool) {
            Write-Log -Message 'Closing runspace pool...' -Level Information
            try { $script:RunspacePool.Close() } catch { Write-Log -Message "Error closing runspace pool: $($_.Exception.Message)" -Level Warning }
            try { $script:RunspacePool.Dispose() } catch { Write-Log -Message "Error disposing runspace pool: $($_.Exception.Message)" -Level Warning }
        }

        # Generate final HTML report
        try {
            Write-Log -Message 'Generating final HTML report...' -Level Information

            $htmlContent = Build-HtmlContent `
                -UserPrintCounts $script:UserPrintCounts `
                -PrinterPrintCounts $script:PrinterPrintCounts `
                -PrintJobs $script:RecentJobs `
                -TemplatePath $script:Config.HtmlTemplatePath

            if (Write-FileWithRetry -Path $script:Config.HtmlFile -Content $htmlContent) {
                Write-Log -Message "Final HTML report generated: $($script:Config.HtmlFile)" -Level Information
            }
        }
        catch {
            # Don't use -ErrorRecord here to avoid recursive error logging
            Write-Warning "Error generating final HTML report: $($_.Exception.Message)"
        }

        Write-Log -Message '=== Print Job Monitor Stopped ===' -Level Information
    }
}

#endregion

#region Script Entry Point

# Start monitoring
try {
    Start-PrintJobMonitoring
}
catch {
    Write-Error "Fatal error in print job monitoring: $($_.Exception.Message)"
    exit 1
}

#endregion
