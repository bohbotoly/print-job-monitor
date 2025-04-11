#region Configuration and Initialization
Import-Module ActiveDirectory
# Configuration: List of servers to monitor
$serverConfigs = @(
    @{
        ServerName = "172.2.0.99" # Changed to your print server
        Description = "Server Description" # Changed description
		
		#You can add more print servers
    }
)
# Path to HTML template file
# Get the directory where the current script is located
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

# Set paths relative to the script location
$htmlTemplatePath = Join-Path -Path $scriptDirectory -ChildPath "template.html"
$outputDirectory = $scriptDirectory
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
}
$dateHTML = Get-Date -Format "dd-MM-yyyy"
$htmlFileBase = Join-Path $outputDirectory "PrintJobsLog-PrintServers-$dateHTML" # Changed filename
$adCache = @{}  
$printerCache = @{} 
$queryInterval = 1  # WMI query polling interval (in seconds) - Lower values increase responsiveness but also load.
$errorCount = 0
$recentJobs = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
$userPrintCounts = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase) # Case-insensitive keys for users
$printerPrintCounts = [System.Collections.Concurrent.ConcurrentDictionary[string,object]]::new([System.StringComparer]::OrdinalIgnoreCase) # Case-insensitive keys for printers
$currentDate = Get-Date -Format "yyyy-MM-dd"
$htmlRefreshInterval = 5  # Seconds between HTML refreshes (reduced for quicker updates)
$lastHtmlRefresh = [datetime]::MinValue
$runspaceCheckInterval = 2 # Seconds between checking runspace status
$lastRunspaceCheck = [datetime]::MinValue
# Verify HTML template file exists
if (-not (Test-Path $htmlTemplatePath)) {
    Write-Error "HTML template file not found at: $htmlTemplatePath"
    exit 1
}
Write-Host "Monitoring Servers:" -ForegroundColor Yellow
$serverConfigs | ForEach-Object { Write-Host "- $($_.ServerName) ($($_.Description))" }
Write-Host "Output Directory: $outputDirectory" -ForegroundColor Yellow
Write-Host "HTML Template: $htmlTemplatePath" -ForegroundColor Yellow
#endregion
#region Utility Functions
function Get-UniqueFileName {
    param (
        [string]$BaseName,
        [string]$Extension
    )
    $counter = 0
    $newFileName = "$BaseName.$Extension"
    while (Test-Path -Path $newFileName) {
        $counter++
        $newFileName = "$BaseName.$($counter.ToString("D2")).$Extension"
    }
    return $newFileName
}
$htmlFile = Get-UniqueFileName -BaseName $htmlFileBase -Extension "html"
Write-Host "HTML Report File: $htmlFile" -ForegroundColor Yellow
# Cache AD User Information
function Get-UserInfo {
    param ([string]$SamAccountName)
    # Normalize the key (e.g., remove domain if present, convert to lower case)
    $normalizedKey = ($SamAccountName -split '\\' | Select-Object -Last 1).ToLowerInvariant()
    if ($adCache.ContainsKey($normalizedKey)) {
        return $adCache[$normalizedKey]
    }
    try {
        # Limit the properties being retrieved
        $user = Get-ADUser -Identity $normalizedKey -Properties DisplayName, Office -ErrorAction Stop
        $userInfo = [PSCustomObject]@{
            DisplayName = $user.DisplayName
            Office = $user.Office
        }
    } catch {
        # Cache failed lookups too, to avoid repeated attempts for non-AD users/errors
        $userInfo = [PSCustomObject]@{
            DisplayName = $SamAccountName # Display original name if lookup fails
            Office = "Unknown"
        }
    }
    # Add to cache with normalized key
    $adCache[$normalizedKey] = $userInfo
    return $userInfo
}
# Improved printer cache
function Get-PrinterName {
    param(
        [string]$PrinterName,
        [string]$ServerName
    )
    # Key format: servername:::printername (lowercase)
    $cacheKey = "$($ServerName):::$($PrinterName)".ToLowerInvariant()
    if ($printerCache.ContainsKey($cacheKey)) {
        return $printerCache[$cacheKey]
    }
    try {
        # Use CIM for potentially better performance/reliability over WMI/Get-Printer ?
        # Or stick with Get-Printer if it works reliably
        $printer = Get-Printer -Name $PrinterName -ComputerName $ServerName -ErrorAction Stop
        $printerInfo = [PSCustomObject]@{
            PrinterName = $printer.Name # Use the name returned by Get-Printer for consistency
        }
        $printerCache[$cacheKey] = $printerInfo
        return $printerInfo
    } catch {
        # Cache failed lookups to avoid retries
        Write-Warning "Could not retrieve details for printer '$PrinterName' on server '$ServerName'. Error: $($_.Exception.Message)"
        $printerInfo = [PSCustomObject]@{
            PrinterName = $PrinterName # Return original name on error
        }
        $printerCache[$cacheKey] = $printerInfo
        return $printerInfo
    }
}
function Build-HTMLContent {
    param (
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$PrintJobs,
        [string]$TemplatePath
    )
    
    # Read the HTML template with explicit UTF8 encoding
    try {
        # Method 1: Using Get-Content with explicit encoding
        $htmlTemplate = Get-Content -Path $TemplatePath -Raw -Encoding UTF8 -ErrorAction Stop
        
        # Alternative Method 2: Using .NET directly if the above doesn't work
        # $htmlTemplate = [System.IO.File]::ReadAllText($TemplatePath, [System.Text.Encoding]::UTF8)
    }
    catch {
        Write-Error "Failed to read HTML template file: $($_.Exception.Message)"
        # Fallback to a minimal template if the file can't be read
        $htmlTemplate = "<html><head><meta charset='UTF-8'></head><body><h1>Error: Could not load template</h1><p>Print monitoring data is still being collected.</p></body></html>"
    }
    
    # Process users, printers, and jobs as before
    # ...
    
    # Batch process the top users
    $topUsersHtml = ""
    # Sort directly on the values within the dictionary entries
    $topUsers = $UserPrintCounts.GetEnumerator() | Sort-Object { $_.Value.TotalPages } -Descending | Select-Object -First 10
    $topUserKey = if ($topUsers.Count -gt 0) { $topUsers[0].Key } else { $null }
    foreach ($userEntry in $topUsers) {
        $userKey = $userEntry.Key
        $userData = $userEntry.Value
        $crownIcon = ""
        if ($userKey -eq $topUserKey) {
            $crownIcon = "👑" # Changed icon
        }
        # Use the cached Get-UserInfo function
        $userInfo = Get-UserInfo -SamAccountName $userKey
        # Fixed the column order to match the HTML table header
        $topUsersHtml += "<tr>
        <td class='highlight' title='$($userInfo.Office)'>$crownIcon $($userInfo.DisplayName)</td>
        <td>$($userData.TotalJobs)</td>
        <td>$($userData.TotalPages)</td>
        </tr>"
    }
    
    # Batch process the top printers
    $topPrintersHtml = ""
    $topPrinters = $PrinterPrintCounts.GetEnumerator() | Sort-Object { $_.Value.TotalPages } -Descending | Select-Object -First 10
    $topPrinterKey = if ($topPrinters.Count -gt 0) { $topPrinters[0].Key } else { $null }
    foreach ($printerEntry in $topPrinters) {
        $printerKey = $printerEntry.Key
        $printerData = $printerEntry.Value
        $crownIcon = ""
        if ($printerKey -eq $topPrinterKey) {
            $crownIcon = "👑" # Changed icon
        }
        # Extract server name and printer name from combined key
        $parts = $printerKey -split ":::"
        $printerNameOnly = $parts[1] # Assuming key is server:::printer
        $serverNameOnly = $parts[0]
        
        # FIX 1: Uppercase the printer names (BSPR)
        # Convert the printer name to uppercase if it starts with bspr
        $displayPrinterName = if ($printerNameOnly -match '^(bspr)') {
            $printerNameOnly.ToUpper()
        } else {
            $printerNameOnly
        }
        
        # Fixed the column order to match the HTML table header
        $topPrintersHtml += "<tr>
        <td class='highlight'>$displayPrinterName $crownIcon</td>
        <td>$($printerData.TotalJobs)</td>
        <td>$($printerData.TotalPages)</td>
        <td>$serverNameOnly</td>
        </tr>"
    }
    
    # Convert queue to array for processing and sort by time with newest first
    $jobArray = $PrintJobs.ToArray() | Sort-Object {
        # Parse the date string into a DateTime object for proper sorting
        if($_.Time -match '(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})') {
            Get-Date -Year $matches[3] -Month $matches[2] -Day $matches[1] -Hour $matches[4] -Minute $matches[5]
        } else {
            [DateTime]::MinValue
        }
    } -Descending
    
    # Create a HashSet to track unique jobs we've already processed
    # This prevents duplicate jobs but allows multiple jobs from same user
    $processedJobs = New-Object System.Collections.Generic.HashSet[string]
    
    # Batch process the print jobs
    $printJobsHtml = ""
    foreach ($job in $jobArray) {
        # Each job needs a unique identifier - either use the JobKey if available or create one
        $uniqueJobId = if ($job.JobKey) { $job.JobKey } else { "$($job.User)-$($job.JobId)-$($job.Time)" }
        
        # Only process this job if we haven't seen it before
        if ($processedJobs.Add($uniqueJobId)) {
            # Use cached Get-UserInfo
            $userInfo = Get-UserInfo -SamAccountName $job.User
            
            # FIX 2: Truncate document name to 40 characters with ellipsis
            $documentName = $job.Document
            if ($documentName.Length -gt 40) {
                $documentName = $documentName.Substring(0, 37) + "..."
                # Store the full document name as a tooltip
                $documentTooltip = "title='$($job.Document)'"
            } else {
                $documentTooltip = ""
            }
            
            # FIX 1: Uppercase printer names for main table too
            $displayPrinterName = if ($job.Printer -match '^(bspr)') {
                $job.Printer.ToUpper()
            } else {
                $job.Printer
            }
            
            # Fixed column order to match HTML table header
            $printJobsHtml += "<tr>
                <td>$($job.Time)</td>
                <td title='$($userInfo.Office)'>$($userInfo.DisplayName)</td>
                <td $documentTooltip>$documentName</td>
                <td>$($job.Pages)</td>
                <td>$displayPrinterName</td>
                <td>$($job.Server)</td>
            </tr>"
        }
    }
    
    # Add a fallback message if no print jobs found
    if ($printJobsHtml -eq "") {
        $printJobsHtml = "<tr><td colspan='6' class='highlight' style='text-align:center;'>לא נמצאו הדפסות ביום הנוכחי</td></tr>"
    }
    
    # Replace placeholders in template with actual data
    $htmlContent = $htmlTemplate -replace '{{topUsersHtml}}', $topUsersHtml
    $htmlContent = $htmlContent -replace '{{topPrintersHtml}}', $topPrintersHtml
    $htmlContent = $htmlContent -replace '{{printJobsHtml}}', $printJobsHtml
    
    return $htmlContent
}
function Safe-WriteToFile {
    param (
        [string]$Path,
        [string]$Content,
        [int]$RetryCount = 3,
        [int]$DelayMilliseconds = 300
    )
    for ($i = 1; $i -le $RetryCount; $i++) {
        try {
            # Ensure directory exists
            $Dir = Split-Path $Path -Parent
            if (-not (Test-Path $Dir)) {
                New-Item -ItemType Directory -Path $Dir -Force | Out-Null
            }
            
            # Use UTF8 encoding with BOM (important for Hebrew text)
            $utf8WithBom = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllText($Path, $Content, $utf8WithBom)
            
            # Alternative method if above doesn't work:
            # $Content | Out-File -FilePath $Path -Encoding UTF8 -Force
            
            return $true # Indicate success
        } catch {
            Write-Warning "Attempt $i : Unable to write to file $Path. Error : $($_.Exception.Message)"
            if ($i -lt $RetryCount) {
                Start-Sleep -Milliseconds $DelayMilliseconds
            }
        }
    }
    Write-Error "Failed to write to file $Path after $RetryCount attempts."
    return $false # Indicate failure
}
#endregion
#region Parallel Monitoring Logic
# Modify the scriptBlock to use a more compatible approach
$scriptBlock = {
    param(
        [string]$ServerName,
        [Parameter(Mandatory=$true)][System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserPrintCounts,
        [Parameter(Mandatory=$true)][System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterPrintCounts,
        [Parameter(Mandatory=$true)][System.Collections.Concurrent.ConcurrentQueue[hashtable]]$RecentJobs,
        [int]$WmiQueryInterval, # Pass interval to the scriptblock
        [hashtable]$SyncHash # Add this parameter for communication with main thread
    )
    # Function to send message back to main thread
    function Send-MessageToMainThread {
        param([string]$Message, [string]$Color = "White")
        
        # Add message to the synchronized queue
        $SyncHash.MessageQueue.Enqueue(@{
            Message = $Message
            Color = $Color
            Timestamp = Get-Date
        })
    }
    
    # Use a boolean flag for cancellation
    $stopRequested = $false
    $SyncHash["StopFlag-$ServerName"] = [ref]$stopRequested
    
    # Register event handler within the runspace for cleanup
    $eventAction = {
        param($Sender, $EventArgs)
        
        # Access the stop flag via synchronized hashtable
        $serverName = $EventArgs.MessageData.ServerName
        $stopFlag = $EventArgs.MessageData.StopFlag
        $sendMessageFunction = $EventArgs.MessageData.SendMessageFunction
        
        # Log the stop request
        & $sendMessageFunction -Message "[$serverName] Stop event received." -Color "Yellow"
        
        # Set the stop flag to true
        $stopFlag.Value = $true
    }
    
    # Register with event args that contain needed data
    $messageData = @{
        ServerName = $ServerName
        StopFlag = [ref]$stopRequested
        SendMessageFunction = ${function:Send-MessageToMainThread}
    }
    
    $null = Register-EngineEvent -SourceIdentifier "StopRunspace-$ServerName" -Action $eventAction -MessageData $messageData
    
    $watcher = $null
    $options = New-Object System.Management.EventWatcherOptions
    # Set a timeout so the loop can check for cancellation
    $options.Timeout = [TimeSpan]::FromSeconds(2)
    
    try {
        Send-MessageToMainThread -Message "[$ServerName] Initializing WMI watcher..." -Color "Cyan"
        $query = "SELECT * FROM __InstanceCreationEvent WITHIN $WmiQueryInterval WHERE TargetInstance ISA 'Win32_PrintJob'"
        $scope = New-Object System.Management.ManagementScope ("\\$ServerName\root\cimv2")
        $scope.Connect() # Explicitly connect
        $watcher = New-Object System.Management.ManagementEventWatcher($scope, $query)
        $watcher.Options = $options
        Send-MessageToMainThread -Message "[$ServerName] Watcher created. Starting event loop..." -Color "Cyan"
        
        # Process print jobs loop
        while (-not $stopRequested) { 
            try {
                # Check for cancellation before waiting for event
                if ($stopRequested) {
                    Send-MessageToMainThread -Message "[$ServerName] Stop requested before WaitForNextEvent." -Color "Yellow"
                    break
                }
                
                # WaitForNextEvent with timeout (set in options above)
                $event = $watcher.WaitForNextEvent()
                
                # Check for cancellation after event wait completes
                if ($stopRequested) {
                    Send-MessageToMainThread -Message "[$ServerName] Stop requested after WaitForNextEvent." -Color "Yellow"
                    if ($event) { $event.Dispose() }
                    break
                }
                
                if ($event) {
                    $job = $event.TargetInstance
                    
                    # Defensive programming: Check if job properties exist
                    $printerName = if ($job.Name) { $job.Name -split "," | Select-Object -First 1 } else { "UnknownPrinter" }
                    $userName = if ($job.Owner) { $job.Owner } else { "UnknownUser" }
                    $documentName = if ($job.Document) { $job.Document } else { "UnknownDocument" }
                    $timeStamp = Get-Date -Format "dd-MM-yyyy HH:mm" # Consistent format
                    
                    # Handle potential 0 pages (e.g., paused jobs initially?) Default to 1.
                    $pageCount = if ($job.TotalPages -and $job.TotalPages -gt 0) { $job.TotalPages } else { 1 }
                    
                    # Get the Job ID for uniqueness
                    $jobId = if ($job.JobId) { $job.JobId } else { [Guid]::NewGuid().ToString() }
                    
                    # Send message to main thread for console display
                    Send-MessageToMainThread -Message "[$ServerName] New Print Job: User=$userName, Printer=$printerName, Pages=$pageCount, Document=$documentName" -Color "Green"
                    
                    # Use lowercase for dictionary keys for consistency
                    $normalizedUserName = ($userName -split '\\' | Select-Object -Last 1).ToLowerInvariant()
                    
                    $printerKey = "$($ServerName):::$($printerName)".ToLowerInvariant()
                    $uniqueId = [Guid]::NewGuid().ToString()
                    $jobKey = "$normalizedUserName-$jobId-$uniqueId"
                    
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
                    
                    # --- Add to Recent Jobs Queue (Thread-Safe) ---
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
                    
                    # Enqueue the new job
                    $RecentJobs.Enqueue($newJobEntry)
                    Send-MessageToMainThread -Message "[$ServerName] Job processed: User='$userName', Printer='$printerName', Pages=$pageCount, JobID=$jobId" -Color "Cyan"
                    $event.Dispose()
                }
            } catch {
                # Check for timeout message which is expected behavior
                if ($_.Exception.Message -like "*Timed out*") {
                    # This is the expected timeout from our 2-second timer
                    # Just continue the loop silently to check for cancellation
                    continue
                } else {
                    Send-MessageToMainThread -Message "[$ServerName] Error during event watch loop: $($_.Exception.Message)" -Color "Red"
                    # Check if we should exit due to stop flag
                    if ($stopRequested) {
                        break
                    }
                    Start-Sleep -Seconds 2 # Pause before retrying after error
                }
            }
        }
    } catch {
        Send-MessageToMainThread -Message "[$ServerName] Failed to initialize WMI watcher. Error: $($_.Exception.Message)" -Color "Red"
        throw $_
    } finally {
        Send-MessageToMainThread -Message "[$ServerName] Runspace is shutting down..." -Color "Yellow"
        if ($watcher) {
            Send-MessageToMainThread -Message "[$ServerName] Stopping and disposing watcher in finally block." -Color "Yellow"
            try { $watcher.Stop() } catch {}
            try { $watcher.Dispose() } catch {}
        }
        Unregister-EngineEvent -SourceIdentifier "StopRunspace-$ServerName" -ErrorAction SilentlyContinue
        Send-MessageToMainThread -Message "[$ServerName] Runspace script block finished." -Color "Yellow"
    }
}

# Function to start or restart monitoring for a single server
function Start-ServerMonitorRunspace {
    param(
        [Parameter(Mandatory=$true)]$ServerConfig,
        [Parameter(Mandatory=$true)]$RunspacePool,
        [Parameter(Mandatory=$true)]$UserPrintCounts,
        [Parameter(Mandatory=$true)]$PrinterPrintCounts,
        [Parameter(Mandatory=$true)]$RecentJobs,
        [int]$WmiQueryInterval,
        [hashtable]$SyncHash 
    )
    $serverName = $ServerConfig.ServerName
    Write-Host "Attempting to start monitoring for server: $serverName" -ForegroundColor Cyan
    try {
        $powershell = [powershell]::Create()
        $null = $powershell.AddScript($scriptBlock).AddParameters(@{
            ServerName         = $serverName
            UserPrintCounts    = $UserPrintCounts 
            PrinterPrintCounts = $PrinterPrintCounts 
            RecentJobs         = $RecentJobs      
            WmiQueryInterval   = $WmiQueryInterval
            SyncHash           = $SyncHash       
        })
        $powershell.RunspacePool = $RunspacePool
        $handle = $powershell.BeginInvoke()
        Write-Host "Successfully initiated monitoring for $serverName." -ForegroundColor Green
        return [PSCustomObject]@{
            PowerShell = $powershell
            Handle     = $handle
            ServerName = $serverName
            StartTime  = Get-Date
            Status     = 'Running'
            LastError  = $null
        }
    } catch {
        Write-Error "Failed to start runspace for $serverName : $($_.Exception.Message)"
        return [PSCustomObject]@{
            PowerShell = $null
            Handle     = $null
            ServerName = $serverName
            StartTime  = Get-Date
            Status     = 'FailedToStart'
            LastError  = $_.Exception.Message
        }
    }
}

# Generate HTML only when needed
function Update-HTMLIfNeeded {
    param (
        [datetime]$LastRefreshTime,
        [int]$RefreshIntervalSeconds,
        [string]$OutputFile,
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$UserCounts,
        [System.Collections.Concurrent.ConcurrentDictionary[string,object]]$PrinterCounts,
        [System.Collections.Concurrent.ConcurrentQueue[hashtable]]$JobsQueue,
        [string]$HtmlTemplatePath
    )
    $currentTime = Get-Date
    if (($currentTime - $LastRefreshTime).TotalSeconds -ge $RefreshIntervalSeconds) {
        Write-Verbose "Refresh interval reached. Generating HTML."
        try {
            $htmlContent = Build-HTMLContent -UserPrintCounts $UserCounts -PrinterPrintCounts $PrinterCounts -PrintJobs $JobsQueue -TemplatePath $HtmlTemplatePath
            if (Safe-WriteToFile -Path $OutputFile -Content $htmlContent) {
                Write-Verbose "HTML file '$OutputFile' updated successfully at $currentTime."
                return $currentTime
            } else {
                 Write-Warning "Failed to write HTML file '$OutputFile' after retries."
                return $LastRefreshTime
            }
        } catch {
            Write-Error "Error generating or writing HTML: $($_.Exception.Message)"
            return $LastRefreshTime
        }
    }
    return $LastRefreshTime
}
#endregion
#region Main Execution Loop
$runspacePool = $null
$activeRunspaces = @{}
# Create a synchronized hashtable for cross-thread communication
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.MessageQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()
try {
    Write-Host "Initializing Runspace Pool..." -ForegroundColor Green
    # Since we only have one server, we just need 1 or 2 threads
    $maxThreads = 2 
    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $maxThreads)
    $runspacePool.Open()
    Write-Host "Starting initial monitoring threads..." -ForegroundColor Green
    foreach ($config in $serverConfigs) {
        $runspaceInfo = Start-ServerMonitorRunspace -ServerConfig $config -RunspacePool $runspacePool -UserPrintCounts $userPrintCounts -PrinterPrintCounts $printerPrintCounts -RecentJobs $recentJobs -WmiQueryInterval $queryInterval -SyncHash $syncHash
        $activeRunspaces[$config.ServerName] = $runspaceInfo
    }
    Write-Host "Monitoring started. Press Ctrl+C to stop." -ForegroundColor Green

    while ($true) {
        while ($syncHash.MessageQueue.Count -gt 0) {
            $message = $null
            if ($syncHash.MessageQueue.TryDequeue([ref]$message)) {
                Write-Host $message.Message -ForegroundColor $message.Color
            }
        }
        
        $today = Get-Date -Format "yyyy-MM-dd"
        if ($today -ne $currentDate) {
            Write-Host "Date changed to $today. Stopping monitoring for log rotation." -ForegroundColor Yellow
            break
        }
        if ((Get-Date) -ge $lastRunspaceCheck.AddSeconds($runspaceCheckInterval)) {
             Write-Verbose "Checking runspace status..."
             $serverNamesToCheck = $activeRunspaces.Keys | Get-Random -Count $activeRunspaces.Count
             foreach ($serverName in $serverNamesToCheck) {
                 $runspaceInfo = $activeRunspaces[$serverName]
                 if ($runspaceInfo.Status -ne 'Running') { continue }
                 if ($runspaceInfo.Handle -and $runspaceInfo.Handle.IsCompleted) {
                    Write-Warning "Runspace for server '$serverName' completed unexpectedly."
                    try {
                        $null = $runspaceInfo.PowerShell.EndInvoke($runspaceInfo.Handle)
                        Write-Warning "Runspace for '$serverName' completed without error? This shouldn't happen with the infinite loop. Restarting."
                        $runspaceInfo.Status = 'CompletedUnexpectedly'
                    } catch {
                        Write-Error "Runspace for server '$serverName' failed. Error: $($_.Exception.Message)"
                        $runspaceInfo.Status = 'Failed'
                        $runspaceInfo.LastError = $_.Exception.Message
                    } finally {
                         try { $runspaceInfo.PowerShell.Dispose() } catch {}
                         $runspaceInfo.PowerShell = $null
                         $runspaceInfo.Handle = $null
                     }
                     Write-Host "Attempting to restart monitoring for $serverName..." -ForegroundColor Yellow
                     $serverConfig = $serverConfigs | Where-Object { $_.ServerName -eq $serverName } | Select-Object -First 1
                     if ($serverConfig) {
                         $newRunspaceInfo = Start-ServerMonitorRunspace -ServerConfig $serverConfig -RunspacePool $runspacePool -UserPrintCounts $userPrintCounts -PrinterPrintCounts $printerPrintCounts -RecentJobs $recentJobs -WmiQueryInterval $queryInterval -SyncHash $syncHash
                         $activeRunspaces[$serverName] = $newRunspaceInfo
                     } else {
                         Write-Warning "Could not find configuration for server $serverName to restart."
                     }
                 } 
             }
             $lastRunspaceCheck = Get-Date
         }
        
        # Update HTML periodically
        $script:lastHtmlRefresh = Update-HTMLIfNeeded -LastRefreshTime $script:lastHtmlRefresh -RefreshIntervalSeconds $htmlRefreshInterval -OutputFile $htmlFile -UserCounts $userPrintCounts -PrinterCounts $printerPrintCounts -JobsQueue $recentJobs -HtmlTemplatePath $htmlTemplatePath
        
        # Small sleep in the main loop to prevent high CPU usage
        Start-Sleep -Milliseconds 100
    }
} finally {
    # --- Cleanup ---
    Write-Host "Stopping monitoring and cleaning up resources..." -ForegroundColor Cyan
    # Stop active runspaces using stop flags
    if ($activeRunspaces) {
        foreach ($serverName in $activeRunspaces.Keys) {
            $runspaceInfo = $activeRunspaces[$serverName]
            if ($runspaceInfo.PowerShell -and $runspaceInfo.Status -eq 'Running') {
                Write-Host "Stopping runspace for $serverName..."
                # Signal the runspace to stop via the event mechanism
                try { 
                    New-Event -SourceIdentifier "StopRunspace-$serverName" -EventArguments @() | Out-Null 
                    
                    # Also directly set the stop flag if possible
                    $stopFlagRef = $syncHash["StopFlag-$serverName"]
                    if ($stopFlagRef) {
                        $stopFlagRef.Value = $true
                        Write-Host "Stop flag for $serverName has been set." -ForegroundColor Cyan
                    }
                } catch {
                    Write-Warning "Error triggering stop event for $serverName : $($_.Exception.Message)"
                }
                
                # Give it a moment to stop gracefully
                Start-Sleep -Seconds 2
                
                # Force stop if still running
                try {
                    if (-not $runspaceInfo.Handle.IsCompleted) {
                        Write-Host "Forcefully stopping runspace for $serverName..." -ForegroundColor Yellow
                        $runspaceInfo.PowerShell.Stop()
                    } else {
                        Write-Host "Runspace for $serverName completed gracefully." -ForegroundColor Green
                    }
                } catch { 
                    Write-Warning "Error stopping PowerShell object for $serverName : $($_.Exception.Message)"
                }
            }
            
            # Clean up and dispose regardless of stop success
            try { 
                if ($runspaceInfo.PowerShell) {
                    if ($runspaceInfo.Handle -and -not $runspaceInfo.Handle.IsCompleted) {
                        # Try to end the invoke if still running
                        try { $null = $runspaceInfo.PowerShell.EndInvoke($runspaceInfo.Handle) } catch {}
                    }
                    $runspaceInfo.PowerShell.Dispose() 
                }
            } catch {
                Write-Warning "Error disposing PowerShell object for $serverName : $($_.Exception.Message)"
            }
        }
    }
    
    # Clean up stop flags
    foreach ($key in @($syncHash.Keys)) {
        if ($key.StartsWith("StopFlag-")) {
            try {
                $syncHash.Remove($key)
            } catch {
                Write-Warning "Error removing stop flag for $key : $($_.Exception.Message)"
            }
        }
    }
    
    # Close the runspace pool
    if ($runspacePool) {
        Write-Host "Closing runspace pool..."
        try { $runspacePool.Close() } catch {Write-Warning "Error closing runspace pool: $($_.Exception.Message)"}
        try { $runspacePool.Dispose()} catch {Write-Warning "Error disposing runspace pool: $($_.Exception.Message)"}
    }
    
    # Generate final HTML output
    try {
        Write-Host "Generating final HTML report..." -ForegroundColor Cyan
        $htmlContent = Build-HTMLContent -UserPrintCounts $userPrintCounts -PrinterPrintCounts $printerPrintCounts -PrintJobs $recentJobs -TemplatePath $htmlTemplatePath
        if (Safe-WriteToFile -Path $htmlFile -Content $htmlContent) {
            Write-Host "Final HTML report generated successfully: $htmlFile" -ForegroundColor Green
        }
    } catch {
        Write-Warning "Error generating final HTML report: $($_.Exception.Message)"
    }
    
    Write-Host "Monitoring stopped." -ForegroundColor Green
}
#endregion