<#
.SYNOPSIS
    Interactive WSUS reporting and maintenance dashboard for Windows Server Update Services.
.DESCRIPTION
    `psc_wsusreporting.ps1` is a menu-driven PowerShell tool that connects to the local WSUS
    server via the Microsoft.UpdateServices.Administration API and provides at-a-glance status,
    on-demand HTML reports, and safe maintenance actions.

    What it shows (live status):
      - Environment: hostname, domain, IP, timestamp
      - Endpoint health: up-to-date, needing updates, with client errors, total clients
      - Update health: installed, needed by computers, with client errors
      - Server stats: approved / declined / pending-approval counts
      - Synchronization: last sync time & result, auto/manual mode, in-progress status & %
      - Content download: in-progress / none, bytes total vs. downloaded
      - Connection info: HTTP/HTTPS, port, WSUS server version
      - Cleanup history: last cleanup export timestamp and suggested next run

    What it can do (menu actions):
      1) Refresh dashboard
      2) Generate **Endpoint Status Report** (HTML)
      3) Generate **Update Status Report** (HTML)
      4) Generate **Last Synchronization Report** (HTML)
      5) **Run WSUS Cleanup** (decline superseded, etc.)
      6) **Test-Run WSUS Cleanup** (no-decline / safe preview)

    Operator experience:
      - Clear, colorized console with a single-key menu
      - Robust try/catch; all steps are timestamp-logged
      - Uses the WSUS Admin API (no PSWindowsUpdate dependency)
      - Stores exports like `SupersededUpdates*.csv` under an `Exports\` folder beside the script
      - `$VersionNumber` banner reflects the current build
.LINK
	https://learn.microsoft.com/en-us/windows/win32/wua_sdk/windows-update-agent--wua--api-reference
    https://learn.microsoft.com/de-de/security-updates/WindowsUpdateServices/18127651
    https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms744624(v=vs.85)
    https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms748969(v=vs.85)
    https://learn.microsoft.com/en-us/previous-versions/windows/desktop/mt748187(v=vs.85)
	https://github.com/PScherling

.NOTES
          FileName: psc_wsusreporting.ps1
          Solution: PSC_WSUS_Reporting
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-05-02
          Modified: 2025-09-25

          Version - 0.0.1 - () - Initial first attempt.
		  Version - 0.0.2 - () - Start-Gen-LastWSUSSynchronizationReport function finished
		  Version - 0.0.3 - () - Start-Gen-EndpointStatusReport function finished
		  Version - 0.0.4 - () - Start-Gen-UpdateStatusReport function WIP
		  Version - 0.0.5 - () - Start-Gen-UpdateStatusReport function finished
		  Version - 0.0.6 - () - Bug fixing
		  Version - 0.0.7 - () - Show in Sync Status; Bug fixing
          Version - 0.1.0 - () - Publishing Version 1.
          Version - 0.1.1 - () - Bug Fixing Sync Progress
		  Version - 0.1.2 - () - Fixing PowerShell syntax issues
		  Version - 0.1.3 - () - Adding WSUS Cleanup functions
		  Version - 0.1.4 - () - Fixed: Minor Issue by HTML Report for Endpoints
		  
		  To-Do:
			- 
			
.REQUIREMENTS
    - Run on the WSUS server with Administrator privileges.
    - PowerShell 5.1+ (or PowerShell 7.x on Windows).
    - WSUS Administration components available (Microsoft.UpdateServices.Administration).
    - Local file system write access for logs/exports.
	- Internet connection to google and microsoft

.OUTPUTS
    - Console dashboard (colorized).
    - Log file: C:\_psc\psc_wsusreporting\Logfiles\psc_wsusreporting.log
    - HTML reports (saved to the script directory or designated report path).
    - CSV exports under: <script folder>\Exports\ (e.g., SupersededUpdates*.csv)
	
.Example
    PS C:\> psc_wsusreporting
    (When the helper module is installed) Starts the tool in a new, maximized PowerShell window.

	PS C:\> powershell.exe -ExecutionPolicy Bypass -File "C:\_psc\psc_wsusreporting\psc_wsusreporting.ps1"
    Launches the WSUS Reporting Tool and opens the interactive dashboard.
#>

# Version number
$VersionNumber = "0.1.3"

# Log file path
$logFile = "C:\_psc\psc_wsusreporting\Logfiles\psc_wsusreporting.log"
if(-not $logfile){
    New-Item -Name "psc_wsusreporting.log" -Path "C:\_psc\psc_wsusreporting\Logfiles" -ItemType "File"
}

# Function to log messages with timestamps
function Write-Log {
	param (
		[string]$Message
	)
	$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
	$logMessage = "[$timestamp] $Message"
	#Write-Output $logMessage
	$logMessage | Out-File -FilePath $logFile -Append
}

# Start logging
Write-Log " Starting wsus-reporting..."
	

####
#### Main Menu
####
function Show-Menu {
	
	Write-Log " Starting Main Menu."
	Write-Log " Gathering system information."
    
	# Function to convert sizes to GB or TB depending on the size
	<#function Convert-Size {
		param (
			[Parameter(Mandatory=$true)]
			[double]$sizeInBytes
		)

		$sizeInGB = [math]::round($sizeInBytes / 1GB, 2)
		if ($sizeInGB -ge 1000) {
			$sizeInTB = [math]::round($sizeInGB / 1024, 2)
			return "$sizeInTB TB"
		} else {
			return "$sizeInGB GB"
		}
	}#>
	
	
	### Variables ###
    
    # System Info
    try{
        $Hostname = $env:COMPUTERNAME
    }
    catch{
        Write-Log "ERROR: Can not fetch Hostname Information. Reason: $_"
		Write-Warning "ERROR: Can not fetch Hostname Information. Reason: $_"
    }

    try{
        $Domain = (Get-WmiObject Win32_ComputerSystem).Domain
    }
    catch{
        Write-Log "ERROR: Can not fetch Domain Information. Reason: $_"
		Write-Warning "ERROR: Can not fetch Domain Information. Reason: $_"
    }

	$global:HTMLReportFileName = ""
    
	# Time and Date Info
    try{
        $Date = Get-Date -UFormat "%A %d/%b/%Y %T"
    }
    catch{
        Write-Log "ERROR: Can not fetch Date Information. Reason: $_"
		Write-Warning "ERROR: Can not fetch Date Information. Reason: $_"
    }
	
	# Network Info
    try{
        $IP = Get-NetIPAddress | Where-Object { $_.InterfaceAlias -notlike "Loopback*" -and $_.PrefixOrigin -notlike "WellKnown" } | Select-Object IPAddress
        $IP = $IP.IPAddress
    }
    catch{
        Write-Log "ERROR: Can not fetch IP Address Information. Reason: $_"
		Write-Warning "ERROR: Can not fetch IP Address Information. Reason: $_"
    }
	
	###
	### Get Last Sync information
	###
	
	# Load WSUS Administration assembly
	try{
		[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
	}
	catch{
		Write-Log "ERROR: Can not fetch WSUS Information. Reason: $_"
		Write-Warning "ERROR: Can not fetch WSUS Information. Reason: $_"
	}

	# Connect to local WSUS server
	Write-Log "Connect to local WSUS serve."
	try{
		$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
	}
	catch{
		Write-Log "ERROR: Can not connect to WSUS. Reason: $_"
		Write-Warning "ERROR: Can not fetch WSUS. Reason: $_"
	}

	# Get last sync history entry
	# Get the synchronization history (use this, it's safe and works)
	Write-Log "Get sync history."
	try{
		$syncHistory = $wsus.GetSubscription().GetSynchronizationHistory() | Sort-Object StartTime -Descending
	}
	catch{
		Write-Log "ERROR: Can not get sync history. Reason: $_"
		Write-Warning "ERROR: Can not get sync history. Reason: $_"
	}

	# Get the most recent sync event
	Write-Log "Get the most recent sync event."
	try{
		$lastSync = $syncHistory | Select-Object -First 1
	}
	catch{
		Write-Log "ERROR: Can not get the most recent sync event. Reason: $_"
		Write-Warning "ERROR: Can not get the most recent sync event. Reason: $_"
	}
	
	if ($null -eq $lastSync -or [string]::IsNullOrWhiteSpace($lastSync.Result)) {
		$lastSyncResult = "No Data available"
	}
	else{
		$lastSyncResult = $lastSync.Result
	}
	
	# filter for a specific report
	#$lastSync = $syncHistory | Where-Object { $_.Id -eq "68bff289-d662-4e6d-aca6-9a3c48524714" }

	# Format the Start and End times
	Write-Log "Format the Start and End times."
	if($lastSync.StartTime -and $lastSync.EndTime){
		$startTimeFormatted = $lastSync.StartTime.ToString("dd MMM. yyyy - HH:mm")
		$endTimeFormatted = $lastSync.EndTime.ToString("dd MMM. yyyy - HH:mm")
	}
	else{
		$startTimeFormatted = "No Data available"
		$endTimeFormatted = "No Data available"
	}
	
	# Check result of last sync
	Write-Log "Check result of last sync."
	if ([string]::IsNullOrWhiteSpace($lastSync.Result)) {
		$lastSyncResult = "No Data available"
	}
	else{
		$lastSyncResult = $lastSync.Result
	}
	
    # Get Endpoints with Errors
	Write-Log "Get Endpoints with Errors."
	if($wsus.GetStatus() | Select-Object ComputerTargetsWithUpdateErrorsCount){
		$EnpointsWithErrors = $wsus.GetStatus() | Select-Object ComputerTargetsWithUpdateErrorsCount
		$EnpointsWithErrors = $EnpointsWithErrors.ComputerTargetsWithUpdateErrorsCount
	}
	else{
		$EnpointsWithErrors = "No Data available"
	}
	
	# Get Endpoints with Pending Updates
	Write-Log "Get Endpoints with Pending Updates."
	if($wsus.GetStatus() | Select-Object ComputerTargetsNeedingUpdatesCount){
		$EnpointsWithPendUpd = $wsus.GetStatus() | Select-Object ComputerTargetsNeedingUpdatesCount
		$EnpointsWithPendUpd = $EnpointsWithPendUpd.ComputerTargetsNeedingUpdatesCount
	}
	else{
		$EnpointsWithPendUpd = "No Data available"
	}
	
	# Get Endpoints with Good State
	Write-Log "Get Endpoints in good state."
	if($wsus.GetStatus() | Select-Object ComputersUpToDateCount){
		$EnpointsGood = $wsus.GetStatus() | Select-Object ComputersUpToDateCount
		$EnpointsGood = $EnpointsGood.ComputersUpToDateCount
	}
	else{
		$EnpointsGood = "No Data available"
	}
	
	# Get Endpoints
	Write-Log "Get Endpoints."
	if($wsus.GetStatus() | Select-Object ComputerTargetCount){
		$EnpointsCount = $wsus.GetStatus() | Select-Object ComputerTargetCount
		$EnpointsCount = $EnpointsCount.ComputerTargetCount
	}
	else{
		$EnpointsCount = "No Data available"
	}
	
	# Get Updates with Errors
	Write-Log "Get Updates with Errors."
	if($wsus.GetStatus() | Select-Object UpdatesWithClientErrorsCount){
		$UpdatesWithErrors = $wsus.GetStatus() | Select-Object UpdatesWithClientErrorsCount
		$UpdatesWithErrors = $UpdatesWithErrors.UpdatesWithClientErrorsCount
	}
	else{
		$UpdatesWithErrors = "No Data available"
	}
	
	# Get Updates needed by Endpoints
	Write-Log "Get Updates needed by Endpoints."
	if($wsus.GetStatus() | Select-Object UpdatesNeededByComputersCount){
		$UpdatesPendEndpoints = $wsus.GetStatus() | Select-Object UpdatesNeededByComputersCount
		$UpdatesPendEndpoints = $UpdatesPendEndpoints.UpdatesNeededByComputersCount
	}
	else{
		$UpdatesPendEndpoints = "No Data available"
	}
	
	# Get Installed Updates
	Write-Log "Get Installed Updates."
	if($wsus.GetStatus() | Select-Object UpdatesUpToDateCount){
		$UpdatesInstalled = $wsus.GetStatus() | Select-Object UpdatesUpToDateCount
		$UpdatesInstalled = $UpdatesInstalled.UpdatesUpToDateCount
	}
	else{
		$UpdatesInstalled = "No Data available"
	}
	
	# Get Updates needing Files
	Write-Log "Get Updates needing Files."
	if($wsus.GetStatus() | Select-Object UpdatesNeedingFilesCount){
		$UpdatesNeedFiles = $wsus.GetStatus() | Select-Object UpdatesNeedingFilesCount
		$UpdatesNeedFiles = $UpdatesNeedFiles.UpdatesNeedingFilesCount
	}
	else{
		$UpdatesNeedFiles = "No Data available"
	}
	
	# Get Updates needing Approval
	Write-Log "Get Updates needing Approval."
	if($wsus.GetStatus() | Select-Object NotApprovedUpdateCount){
		$UpdatesPendApproval = $wsus.GetStatus() | Select-Object NotApprovedUpdateCount
		$UpdatesPendApproval = $UpdatesPendApproval.NotApprovedUpdateCount
	}
	else{
		$UpdatesPendApproval = "No Data available"
	}
	
	# Get Approved Updates
	Write-Log "Get Approved Updates."
	if($wsus.GetStatus() | Select-Object ApprovedUpdateCount){
		$UpdatesApprovedCount = $wsus.GetStatus() | Select-Object ApprovedUpdateCount
		$UpdatesApprovedCount = $UpdatesApprovedCount.ApprovedUpdateCount
	}
	else{
		$UpdatesApprovedCount = "No Data available"
	}
	
	# Get Declined Updates
	Write-Log "Get Declined Updates."
	if($wsus.GetStatus() | Select-Object DeclinedUpdateCount){
		$UpdatesDeclinedCount = $wsus.GetStatus() | Select-Object DeclinedUpdateCount
		$UpdatesDeclinedCount = $UpdatesDeclinedCount.DeclinedUpdateCount
	}
	else{
		$UpdatesDeclinedCount = "No Data available"
	}
	
	# Get Connection Type
	Write-Log "Get Connection Type."
	if($wsus.UseSecureConnection -eq $true){
		$ConnectionType = "Secure Connection"
	}
	else{
		$ConnectionType = "Not Secured"
	}
	
	# Get last sync type
	Write-Log "Get last sync type."
	if($wsus.GetSubscription() | select-object SynchronizeAutomatically){
		$LastSyncType = $wsus.GetSubscription() | select-object SynchronizeAutomatically
		$LastSyncType = $LastSyncType.SynchronizeAutomatically
		
		if($LastSyncType -eq $true){
			$LastSyncType = "Automatically"
		}
		else{
			$LastSyncType = "Manual"
		}
	}
	else{
		$ConnectionType = "Not Secured"
	}
	
	# Check for in Progress File Download
	Write-Log "Check for in Progress File Download."
	if($wsus.GetContentDownloadProgress() | select-object *){
		$TotalBytes = $wsus.GetContentDownloadProgress() | select-object TotalBytesToDownload
		$TotalBytes = $TotalBytes.TotalBytesToDownload
		
		$DownloadedBytes = $wsus.GetContentDownloadProgress() | select-object DownloadedBytes
		$DownloadedBytes = $DownloadedBytes.DownloadedBytes
		
		Write-Log "File Download:"
		Write-Log "    $TotalBytes"
		Write-Log "    $DownloadedBytes"
		
		if($TotalBytes -eq "0" -or $DownloadedBytes -eq $null){
			$FileDownloadProgress = "No files downloaded yet"
		}
		elseif($TotalBytes -eq $DownloadedBytes){
			$FileDownloadProgress = "No file download in progress"
		}
		elseif($TotalBytes -gt $DownloadedBytes){
			$FileDownloadProgress = "File download in progress"
		}
		else{
			$FileDownloadProgress = "No data vailable"
		}
	}
	else{
		$FileDownloadProgress = "No data vailable"
	}
	
	# Check for in Progress Synchronization
	try {
		$subscription = $wsus.GetSubscription()
		
		$syncStatus = $subscription.GetSynchronizationStatus()
		$syncProgress = $subscription.GetSynchronizationProgress()
    
		

		if ($syncStatus -eq "Running") {
			$syncStatus = "Synchronization is currently in progress"
		} else {
			$syncStatus = "No synchronization is currently in progress"
		}
		
		if($syncProgress.TotalItems -ne 0 -and $syncProgress.ProcessedItems -ne 0){
			$percentage = ($syncProgress.ProcessedItems / $syncProgress.TotalItems) * 100
			$percentage = "{0:N1}%" -f $percentage
		}
        elseif($syncStatus -eq "Running" -and $syncProgress.ProcessedItems -eq 0) {
            $percentage = "Collecting Data"
        }
		else{
			$percentage = "None"
		}
	}
	catch {
		Write-Warning "ERROR: Unable to determine synchronization status. Reason: $_"
	}

	# Get WSUS Cleanup Info
	$cleanupPath = Split-Path $script:MyInvocation.MyCommand.Path
	$SupersededList = Join-Path $cleanupPath "Exports\SupersededUpdates*.csv"
	try{
		$GetSupersededFile = Get-ChildItem -Path $SupersededList |
    		Sort-Object LastWriteTime -Descending |
    		Select-Object -First 1
	}
	catch{
		Write-Log "ERROR: SupersededList from Cleanup not found. Reason: $_"
	}
	if($GetSupersededFile) {
		$modifiedDate = (Get-Item "$($GetSupersededFile)").LastWriteTime
		$modifiedDateFormatted = $modifiedDate.ToString("yyyy-MM-dd HH:mm:ss")

		$sugCleanup = $modifiedDate.AddDays(365)
		$sugCleanupFormatted = $sugCleanup.ToString("yyyy-MM-dd")
	}
	else{
		$modifiedDateFormatted = "No Data"
		$sugCleanupFormatted = "No Data"
	}
	
	Write-Log "Endpoints with Errors: $EnpointsWithErrors"
	Write-Log "Endpoints with pending updates: $EnpointsWithPendUpd"
	Write-Log "Good Endpoints: $EnpointsGood"
	Write-Log "Updates with Errors: $UpdatesWithErrors"
    Write-Log "Updates required by endpoints: $UpdatesPendEndpoints"
    Write-Log "Installed Updates: $UpdatesInstalled"
	Write-Log "Updates with Errors: $UpdatesWithErrors"
    Write-Log "Updates required by endpoints: $UpdatesPendEndpoints"
    Write-Log "Installed Updates: $UpdatesInstalled"
	Write-Log "Updates waiting for approval: $UpdatesPendApproval"
    Write-Log "Approved updates: $UpdatesApprovedCount"
    Write-Log "Declined updates: $UpdatesDeclinedCount"
    Write-Log "Registered Endpoints: $EnpointsCount"
    Write-Log "Status: $FileDownloadProgress"
    Write-Log "File download status: $UpdatesNeedFiles Updates needing Files"
	Write-Log "Sync Status: $syncStatus"
	Write-Log "Sync Progress: $percentage"
	Write-Log "Last synchronization type: $LastSyncType"
    Write-Log "Last Synchronization: $endTimeFormatted"
    Write-Log "Last synchronization result: $lastSyncResult"
    Write-Log "Type: $ConnectionType"
    Write-Log "Port: $($wsus.PortNumber)"
    Write-Log "Server Version: $($wsus.Version)"
	Write-Log "Cleanup executed: $modifiedDateFormatted"
	Write-Log "Suggested Cleanup execute: $sugCleanupFormatted"
	
	
    ###
    ### Showing the menu
    ###
    Clear-Host

	Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"
    Write-Host "-----------------------------------------------------------------------------------"
    Write-Host "              System Information"
    Write-Host "-----------------------------------------------------------------------------------"
    Write-Host "
    + Version                                                $VersionNumber
    + $Date

    + Domain/Workgroup                                       $Domain
    + Hostname                                               $Hostname
    + IP Address                                             $IP
	
    + Endpoint Status
        Endpoints with Errors:                               $EnpointsWithErrors
        Endpoints with pending updates:                      $EnpointsWithPendUpd
        Good Endpoints:                                      $EnpointsGood
		
    + Update Status
        Updates with Errors:                                 $UpdatesWithErrors
        Updates required by endpoints:                       $UpdatesPendEndpoints
        Installed Updates:                                   $UpdatesInstalled
	
    + Server Statistic
        Updates waiting for approval:                        $UpdatesPendApproval
        Approved updates:                                    $UpdatesApprovedCount
        Declined updates:                                    $UpdatesDeclinedCount
        Registered Endpoints:                                $EnpointsCount
	
    + Synchronization Status
        Download Status:                                     $FileDownloadProgress
        File download status:                                $UpdatesNeedFiles Updates needing Files
        Sync Status:                                         $syncStatus
        Sync Progress:                                       $percentage
        Last synchronization type:                           $LastSyncType
        Last synchronization:                                $endTimeFormatted
        Last synchronization result:                         $lastSyncResult
	
    + Connection Information
        Type:                                                $ConnectionType
        Port:                                                $($wsus.PortNumber)
        Server Version:                                      $($wsus.Version)

    + Cleanup Information
        Last run:                                            $modifiedDateFormatted
        Suggested next run:                                  $sugCleanupFormatted
	"
      
    Write-Host "`n-----------------------------------------------------------------------------------"
    Write-Host "               Windows Update Services Reporting Tool"
    Write-Host "-----------------------------------------------------------------------------------"
    
    Write-Host "
    Actions
    1) Refresh
    2) Endpoints Status Report
    3) Update Status Report
    4) Last Synchronization Report
    5) Run WSUS Cleanup
    6) Test-Run WSUS Cleanup
	
    e) Leave"

    
	do {
		$choice = Read-Host " Choose an Option (1-6/e)"
		Write-Log " User Input: $choice"
		switch ($choice) {
			1 { Show-Menu }
			2 { Start-Gen-EndpointStatusReport }
			3 { Start-Gen-UpdateStatusReport }
			4 { Start-Gen-LastWSUSSynchronizationReport }
			5 { Start-Run-Cleanup }
			6 { Start-Run-Cleanup -SkipDecline }
			e { Exit }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne {1..6} -and $choice -ne "e")
}


###
### WSUS Cleanup
###
function Start-Run-Cleanup {
	[CmdletBinding()]
	Param(
		[switch] $SkipDecline
	)

	# Report Version
	$CleanupVersion = "0.1"

	$CompName = $env:COMPUTERNAME
	$DateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
	
	# Start logging	
	Write-Log "Start 'WSUS Cleanup Report' Logging."

	$FileName = $Hostname+"_WU-Cleanup-Report_"+$DateTime+".html"
	$RootDir = "C:\_psc\psc_wsusreporting\"
	$FileDir = "C:\_psc\psc_wsusreporting\Exports\"
	$MediaDir = "C:\_psc\psc_wsusreporting\Exports\Media\"
	$FilePath = "$FileDir$FileName"

	$HTMLReportFileSelected = "false"
	$GetHTMLReportFiles = ""
	$SumOfHTMLReportFile = ""
	$HTMLReportFile = ""
	$HTMLReportFileList = @()
	
	$outSupersededList = Join-Path $FileDir "SupersededUpdates_$($DateTime).csv"
	$outSupersededListBackup = Join-Path $FileDir "SupersededUpdatesBackup_$($DateTime).csv"
	
	#$modifiedDate = (Get-Item "$($outSupersededList)").LastWriteTime
	#Write-Host "Modified on: $modifiedDate"

	$countAllUpdates = 0
	$countSupersededAll = 0
	$countSupersededLastLevel = 0
	$countSupersededExclusionPeriod = 0
	$countSupersededLastLevelExclusionPeriod = 0
	$countDeclined = 0

	function Start-WSUS-Cleanup 
	{
		# Function to extract Details of Update
		function Get-UpdateDetails($update) 
		{
			if($update.IsSuperseded){
				$IsSuperseded = "Superseded"
			}
			else{
				$IsSuperseded = "-"
			}

			if($update.HasSupersededUpdates){
				$HasSuperseded = "Yes"
			}
			else{
				$HasSuperseded = "No"
			}
			
			$kb = ($update.KnowledgebaseArticles -join ", ")
			if($kb){
				$kb = "KB$($kb)"
			}
			else{
				$kb = "-"
			}
			$url = $update.AdditionalInformationUrls -join ", "
			$product = $update.ProductTitles -join ", "
			$gID = $update.Id.UpdateId.Guid.ToString()

			return [PSCustomObject]@{
				UpdateID 			= $gID
				ProductTitles 		= $product
				Classification 		= $update.UpdateClassificationTitle
				KBArticle   		= $kb
				Title       		= $update.Title
				SupportURL  		= $url
				HasSuperseded		= $HasSuperseded
			}

		}

		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 10
		Write-Log "Fetch WSUS Report Information."
		# Load WSUS Administration assembly
		try{
			[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
		}
		catch{
			Write-Log "ERROR: Can not fetch WSUS Information. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS Information. Reason: $_"
		}

		# Connect to local WSUS server
		Write-Log "Connect to local WSUS server."
		try{
			$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
		}
		catch{
			Write-Log "ERROR: Can not connect to WSUS. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS. Reason: $_"
		}

		#Write-Host "Getting a list of all updates... " -NoNewLine
		Write-Log "Getting a list of all updates."
		try {
			$allUpdates = $wsus.GetUpdates()
			$countAllUpdates = $allUpdates.Count
		}
		catch
		{
			Write-Log "ERROR: Failed to get updates. Reason: $_"
			Write-Log "If this operation timed out, please decline the superseded updates from the WSUS Console manually."
			Write-Warning "ERROR: Failed to get updates. Reason: $_
    If this operation timed out, please decline the superseded updates from the WSUS Console manually."

			return
		}

		# Format the Start time
		Write-Log "Format the Start time."
		$startTimeFormatted = Get-Date -Format "dd MMM. yyyy - HH:mm"

		# Format Report creation time
		Write-Log "Format Report creation time."
		$ReportTime = Get-Date -Format "dd MMM. yyyy - HH:mm"

		Write-Log "Parsing the list of updates."
		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 30

		
		$supersededUpdates = $allUpdates | Where-Object { $_.IsSuperseded -eq $true -and $_.IsDeclined -eq $false }
		$countSupersededAll = $supersededUpdates.Count

		$declinedUpdates = $allUpdates | Where-Object { $_.IsDeclined -eq $true }
		$countDeclined = $declinedUpdates.Count

	<#
		foreach($update in $allUpdates) {
		
			$countAllUpdates++
			
			if ($update.IsDeclined) {
				$countDeclined++
			}
			

			if (!$update.IsDeclined -and $update.IsSuperseded) {
				$countSupersededAll++
				
				if (!$update.HasSupersededUpdates) {
					$countSupersededLastLevel++
				}

				if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
					$countSupersededExclusionPeriod++
					if (!$update.HasSupersededUpdates) {
						$countSupersededLastLevelExclusionPeriod++
					}
				}		
				
				"$($update.Id.UpdateId.Guid), $($update.Id.RevisionNumber), $($update.Title), $($update.KnowledgeBaseArticles), $($update.SecurityBulletins), $($update.HasSupersededUpdates)" | Out-File $outSupersededList -Append       
				
			}
		}
	#>
		
		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 40
		$i = 0
		if (!$SkipDecline) {
			
			Write-Log "SkipDecline flag is set to $SkipDecline. Continuing with declining updates"
			$updatesDeclined = 0
			
			if ($DeclineLastLevelOnly) {
			<#
				Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 50
				Write-Log "DeclineLastLevel is set to True. Only declining last level superseded updates." 
				
				foreach ($update in $allUpdates) {
					
					if (!$update.IsDeclined -and $update.IsSuperseded -and !$update.HasSupersededUpdates) {
						if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  {
							$i++
							$percentRaw = ($updatesDeclined / $countSupersededLastLevelExclusionPeriod) * 100
							$percentComplete = [math]::Round($percentRaw)

							Write-Progress -id 2 -Activity "Declining Updates" -Status "Declining update #$i/$countSupersededLastLevelExclusionPeriod - $($update.Id.UpdateId.Guid)" -PercentComplete $percentComplete -CurrentOperation "$($percentComplete)% complete"
							
							try 
							{
								$update.Decline()                    
								$updatesDeclined++
							}
							catch
							{
								Write-Log "WARNING: Failed to decline update $($update.Id.UpdateId.Guid). Reason: $_"
								Write-Warning: "Failed to decline update $($update.Id.UpdateId.Guid). Reason: $_"
							} 
						}             
					}
				}   
			#>
			}
			else {
				Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 50
				Write-Log "DeclineLastLevel is set to False. Declining all superseded updates."
				
				foreach ($update in $supersededUpdates) #($update in $allUpdates) 
				{
					
					#if (!$update.IsDeclined -and $update.IsSuperseded) 
					#{
					if ($update.CreationDate -lt (get-date).AddDays(-$ExclusionPeriod))  
					{   
						
						$i++							
						$percentRaw = ($updatesDeclined / $countSupersededAll) * 100
						$percentComplete = [math]::Round($percentRaw)
						
						Write-Progress -Activity "Declining Updates" -Status "Declining update #$i/$countSupersededAll - $($update.Id.UpdateId.Guid)" -PercentComplete $percentComplete -CurrentOperation "$($percentComplete)% complete"
						try 
						{
							$update.Decline()
							$updatesDeclined++
						}
						catch
						{
							Write-Log "WARNING: Failed to decline update $($update.Id.UpdateId.Guid). Reason: $_"
							Write-Warning "Failed to decline update $($update.Id.UpdateId.Guid). Reason: $_"
						}
					}
					#}
				}   
				
			}
			Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 60
			Write-Log "  Declined $updatesDeclined updates."
			if ($updatesDeclined -ne 0) {
				Copy-Item -Path $outSupersededList -Destination $outSupersededListBackup -Force
				Write-Host "Backed up list of superseded updates to $outSupersededListBackup"
				Write-Log "Backed up list of superseded updates to $outSupersededListBackup"
			}
			Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 70
		}
		else {
			#Write-Host "SkipDecline flag is set to $SkipDecline. Skipped declining updates"
			Write-Log "SkipDecline flag is set to $SkipDecline. Skipped declining updates"
		}
		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 80
		
		# Format End Time
		Write-Log "Format the End time."
		$endTimeFormatted = Get-Date -Format "dd MMM. yyyy - HH:mm"

		# Get Details
		$supersededUpdates = $supersededUpdates | ForEach-Object { 
            Get-UpdateDetails $_
        } | Sort-Object -Property Title -Descending

		<#
		# Process and output
		#>
		Add-Content $FilePath -Value @"
    <h2>Report Information</h2>		
    <table class="report-info">
	<tbody>
	<tr>
	<td>Cleanup Start Time:</td>
	<td>$startTimeFormatted</td>
	</tr>
	<tr>
	<td>Cleanup End Time:</td>
	<td>$endTimeFormatted</td>
	</tr>
	<tr>
	<td>Report generated:</td>
	<td>$ReportTime</td>
	</tr>
    <tr>
    <td>WSUS Server:</td>
    <td>$CompName</td>
    </tr>
	<tr>
	<td>SkipDecline Flag:</td>
	<td>$SkipDecline</td>
	</tr>
	<tr>
	<td>Number of Updates:</td>
	<td>$countAllUpdates</td>
	</tr>
	<tr>
	<td>Number of Superseded Updates:</td>
	<td>$countSupersededAll</td>
	</tr>
    </tbody>
    </table>
"@


		#Create Table for New Updates
		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 90
		Add-Content $FilePath -Value @"
    <h2 id="superseded-updates">Superseded Updates</h2>
    <table class="list-syncedUpdates">
    <tbody>
    <tr>
	<th>UpdateID</th>
    <th>Product Titles</th>
    <th>Classification</th>
	<th>KBArticle</th>
	<th>Title</th>
	<th>HasSuperseded</th>
	<th>SupportURL</th>
    </tr>
"@

		$UpdateIDList = @()
		$UpdateProductTitleList = @()
		$UpdateClassificationList = @()
		$UpdateKBArticleList = @()
		$UpdateTileList = @()
		$UpdateSupportURLList = @()
		$UpdateHasSupersededUpdatesList = @()

		#Fill Table with content
		if($supersededUpdates){
			foreach ($SupersededUpdate in $supersededUpdates) {
				$UpdateIDList += $SupersededUpdate.UpdateID
				$UpdateProductTitleList += $SupersededUpdate.ProductTitles
				$UpdateClassificationList += $SupersededUpdate.Classification
				$UpdateKBArticleList += $SupersededUpdate.KBArticle
				$UpdateTileList += $SupersededUpdate.Title
				$UpdateSupportURLList += $SupersededUpdate.SupportURL
				$UpdateHasSupersededUpdatesList += $SupersededUpdate.HasSuperseded
			}

			$supersededUpdates |  
				select-object UpdateID,ProductTitles,Classification,KBArticle,Title,HasSuperseded,SupportURL |
				sort-object Title -Descending |
				Export-Csv -Path $outSupersededList -NoTypeInformation -Encoding UTF8 -Delimiter ";"

			for(($x = 0); $x -lt $supersededUpdates.Count; $x++) {
				$UpdateID = $UpdateIDList[$x]
				$ProductTitle = $UpdateProductTitleList[$x]
				$Classification = $UpdateClassificationList[$x]
				$KBArticle = $UpdateKBArticleList[$x]
				$UpdateTitle = $UpdateTileList[$x]
				$HasSuperseded = $UpdateHasSupersededUpdatesList[$x]
				$SupportURL = $UpdateSupportURLList[$x]

				Add-Content $FilePath -Value @"
			<tr>
			<td>$UpdateID</td>
			<td>$ProductTitle</td>
			<td>$Classification</td>
			<td>$KBArticle</td>
			<td>$UpdateTitle</td>
			<td>$HasSuperseded</td>
			<td class="Url"><a href="$SupportURL" target="_blank">$SupportURL</a></td>
			</tr>
"@
			}

		}
		else{
			Add-Content $FilePath -Value @"
			<tr>
			<td>No Data available</td>
			</tr>
"@
		}
		
		#Finish Table
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@
		
		### Print Output to Console ###
		Write-Log "Report Version                 $CleanupVersion"
		Write-Log "Cleanup Start Time:            $startTimeFormatted"
		Write-Log "Cleanup End Time:              $endTimeFormatted"
		Write-Log "Report generated:              $ReportTime"
		Write-Log "WSUS Server:                   $CompName"
		Write-Log "SkipDecline flag:              $SkipDecline"
		Write-Log "All Updates:                   $countAllUpdates"
		$exDclUpd = ($countAllUpdates - $countDeclined)
		Write-Log "Any except Declined:           $exDclUpd"
		Write-Log "All Superseded Updates:        $countSupersededAll"
		$intUpd = ($countSupersededAll - $countSupersededLastLevel)
		Write-Log "Superseded Updates (Intermediate): $intUpd"
		Write-Log "Superseded Updates (Last Level): $countSupersededLastLevel"
		Write-Log "List of superseded Updates: $outSupersededList"
		
		Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"

		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "              Cleanup Information"
		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "
    + Report Version           $CleanupVersion

    + Cleanup Start Time:      $startTimeFormatted
    + Cleanup End Time:        $endTimeFormatted
    + Report generated:        $ReportTime
    + WSUS Server:             $CompName
    + SkipDecline flag:        $SkipDecline
                        
    + All Updates:             $($countAllUpdates)
    + Any except Declined:     $($exDclUpd)
    + All Superseded Updates:  $($countSupersededAll)

    + Superseded Updates (Intermediate):    $intUpd
    + Superseded Updates (Last Level):      $countSupersededLastLevel

    + List of superseded Updates:           $outSupersededList
"

		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Yellow "              Superseded Updates:"
		#Write-Host "-----------------------------------------------------------------------------------"
		if($supersededUpdates){
			$supersededUpdates | Sort-Object -Property Title -Descending | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		}

		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 100
	}

	function SelectHTMLReportFile {	
		$Options = @()
		$Input = ""
		$HTMLReportFileList = @()  
		$HTMLReportFileSelected = $false

		$GetHTMLReportFiles = Get-ChildItem -File "$FileDir" -Name -Include *.html
		$SumOfHTMLReportFile = $GetHTMLReportFiles.Count

		while (-not $HTMLReportFileSelected) {
			foreach ($HTMLReportFile in $GetHTMLReportFiles) {
				$HTMLReportFileList += $HTMLReportFile
			}

			Write-Host "`n Select a Report File:"
			for (($x = 0); $x -lt $SumOfHTMLReportFile; $x++) {
				Write-Host "    $x) $($HTMLReportFileList[$x])"
				$Options += $x
			}

			$Input = Read-Host "`n Please select an option (0/1/..)"
			Write-Log " User Input: $Input"

			if ($Options -contains [int]$Input) {
				$HTMLReportFileSelected = $true
				$GetFile = $HTMLReportFileList[$Input]
				$global:HTMLReportFileName = $GetFile
				Write-Host "`n Selected HTML Report File: $global:HTMLReportFileName" -ForegroundColor Yellow
			}
			else {
				Write-Host "`n Wrong Input" -ForegroundColor Red
				Write-Log " Wrong Input."
			}
		}
	}

	function ConvertToPDF {

		#$reportname = Get-ChildItem -Path C:\_psc\WSUS Files\WSUS_Reporting\Exports\*.html -Name
		#$reportname = $global:HTMLReportFileName
		$global:pdfPath = "$FileDir$htmlfilename.pdf"
		
		$msedge = @(
			"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
			"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
		)
		$msedge = $msedge | Where-Object { Test-Path $_ } | Select-Object -First 1
		
		if (-not $msedge) {
			Write-Log "ERROR: Microsoft Edge not found. Please install Microsoft Edge."
			Write-Host -ForegroundColor Red "Microsoft Edge not found. Please install Microsoft Edge."
			return
		}
		else{
		
			Write-Log "Converting HTML Report to PDF"
			Write-Log "Selected File: $htmlPath"
			Write-Log "Converted File: $global:pdfPath"

			Start-Sleep 5
		
			$arguments = "--print-to-pdf=""$global:pdfPath"" --headless --page-size=A4 --disable-extensions --no-pdf-header-footer --disable-popup-blocking --run-all-compositor-stages-before-draw --disable-checker-imaging --landscape ""file:///$htmlPath"""
			Write-Log "Edge arguments: $arguments"
			
			try{
			
				#Start-Process "msedge.exe" -ArgumentList @("--headless","--print-to-pdf=""$global:pdfPath""", "--landscape", "--page-size=A4","--disable-extensions","--no-pdf-header-footer","--disable-popup-blocking","--run-all-compositor-stages-before-draw","--disable-checker-imaging", "file:///$htmlPath") -Wait
				Start-Process -FilePath "$msedge" -ArgumentList $arguments -Wait -NoNewWindow
			
			}
			catch{
				Write-Log "ERROR: Convering to pdf failed. Reason: $_"
				Write-Host -ForegroundColor yellow "ERROR: Convering to pdf failed. Reason: $_"
			}
			Start-Sleep 5
		}
	}
	
	function GenerateHTMLReport 
	{
    
		#Wait 10 Seconds - System needs to start background services etc. after foregoing reboot.
		Write-Progress -id 1 -Activity "Generating WSUS Cleanup Report" -Status "Generating Report:" -PercentComplete 0
		Start-Sleep -Seconds 10

		If (-Not ( Test-Path $RootDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'psc_wsusreporting'."
			New-Item -Path "C:\_psc\" -Name "psc_wsusreporting" -ItemType "directory"
		}

		If (-Not ( Test-Path $FileDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Exports'."
			New-Item -Path $RootDir -Name "Exports" -ItemType "directory"
		}

		If (-Not ( Test-Path $MediaDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Media'."
			New-Item -Path $FileDir -Name "Media" -ItemType "directory"
		}

		#Create HTML Report
		Write-Log "Create HTML Report."
		If (-Not ( Test-Path $FilePath ))
		{
			
			#Copy CSS Stylesheet and Images from DeploymentShare
			#Write-Log "Copy CSS Stylesheet and Images from DeploymentShare."
			#Copy-Item -Path "$global:FileShare\*" -Destination $global:MediaDir -Recurse -Force
			
			#Create File
			Write-Log "Create File."
			New-Item $FilePath -ItemType "file" | out-null
		
			#Add Content to the File
			Write-Log "Add Content to the File."
			Add-Content $FilePath -Value @"
<!doctype html>
<html>
	<head>
		<title>WSUS Synchronization Report for $Hostname</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="description" content="this is a Report of the synchronization of windows updates."/>
		<meta name="thumbnail" content=""/>
		<link rel="stylesheet" href="Media/styles.css" />
	</head>
	<body>
	<div id="main">
		<div id="title">
			<img id=""default_logo" src="Media/Powershell_logo.png" alt="Logo">
			<h1 id="title">WSUS Synchronization Report for $Hostname</h1>
            <table class="report-info">
			<tbody>
			<tr>
			<td>Report Template Version:</td>
			<td>Version $CleanupVersion</td>
			</tr>
			</tbody>
			</table>
        </div>
"@
		}

		###### Function Calls ######
		Start-WSUS-Cleanup

		#Finish HTML Report
		Add-Content $FilePath -Value @"
</div>
</body>
<footer>
</footer>
</html>
"@

	}

	function Start-Run-DeleteSupersededUpdates 
	{
		Write-Host "    Starting to delete superseded updates."
		Write-Log "Starting to delete superseded updates."
		Write-Log "Cancelling all downloads."
		try{
			(Get-WsusServer).CancelAllDownloads()
		}
		catch{
			Write-Log "ERROR: Downloads could not be canceled. Reason: $_"
			Write-Warning "ERROR: Downloads could not be canceled. Reason: $_"
		}

		Write-Host "    Stopping relevant services."
		Write-Log "Stopping relevant services."
		try{
			Stop-Service -Name WsusService,BITS -Force #-WhatIf
		}
		catch{
			Write-Log "ERROR: Services could not be stopped. Reason: $_"
			Write-Warning "ERROR: Services could not be stopped. Reason: $_"
		}

		Write-Host "    Deleting temporary directories."
		Write-Log "Deleting temporary directories."
		try{
			Remove-Item -Path $env:LOCALAPPDATA\Temp\* -Recurse -ErrorAction SilentlyContinue #-WhatIf
			Remove-Item -Path $env:SystemRoot\Temp\* -Recurse -ErrorAction SilentlyContinue #-WhatIf
		}
		catch{
			Write-Log "ERROR: Temporary directories could not be deleted. Reason: $_"
			Write-Warning "ERROR: Temporary directories could not be deleted. Reason: $_"
		}

		Write-Host "    Starting relevant services."
		Write-Log "Starting relevant services."
		try{
			Start-Service -Name WsusService,BITS #-WhatIf
		}
		catch{
			Write-Log "ERROR: Services could not be started. Reason: $_"
			Write-Warning "ERROR: Services could not be started. Reason: $_"
		}

		Write-Host "    Get Update server again."
		Write-Log "Get Update server again."
		try{
			[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
			$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer();
		}
		catch{
			Write-Log "ERROR: Could not load WSUS Services. Reason: $_"
			Write-Warning "ERROR: Could not load WSUS Services. Reason: $_"
		}

		Write-Host "    Delete Updates."
		Write-Log "Delete Updates."
		try{
			$declined = $wsus.GetUpdates() | Where {$_.IsDeclined -eq $true}
			$declined | ForEach-Object { $wsus.DeleteUpdate($_.Id.UpdateId.ToString()); Write-Host $_.Title removed; $_.Title }
		}
		catch{
			Write-Log "ERROR: Could not load WSUS Services. Reason: $_"
			Write-Warning "ERROR: Could not load WSUS Services. Reason: $_"
		}

		Write-Host "    Start WSUS Cleanup Tool."
		Write-Log "Start WSUS Cleanup Tool."
		try{
			Get-WsusServer |
				Invoke-WsusServerCleanup `
					-CleanupObsoleteComputers `
					-CleanupObsoleteUpdates `
					-CleanupUnneededContentFiles `
					-CompressUpdates `
					-DeclineExpiredUpdates `
					-DeclineSupersededUpdates `
					-Verbose #`
					#-WhatIf
		}
		catch{
			Write-Log "ERROR: Could not load WSUS Cleanup Tool. Reason: $_"
			Write-Warning "ERROR: Could not load WSUS Cleanup Tool. Reason: $_"
		}

		Write-Host "
    Actions:

    e) Return to main menu"

    
		do {
			$choice = Read-Host " Choose an Option"
			Write-Log " User Input: $choice"
			switch ($choice) {
				e { Show-Menu }
				default { 
					Write-Log " Wrong Input."
					Write-Host "Wrong Input. Please choose an option above." 
				}
			}

		} while ($choice -ne "e")
		
		# Finish logging	
		Write-Log "Finish 'WSUS Cleanup Report' Logging."
	}

	<#
	# Starting Script
	#>
	Clear-Host
	GenerateHTMLReport
	SelectHTMLReportFile
	# Get HTML Report information
	$htmlfilename = $global:HTMLReportFileName -split ".html" | select-object -First 1
	$htmlPath = "$FileDir$global:HTMLReportFileName"
		
	
	do {
		$choice = Read-Host "`n Do you want to convert a report to pdf? (y/n)"
		Write-Log " User Input: $choice"
		switch ($choice) {
			"y" { ConvertToPDF }
			"n" {  }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne "y" -and $choice -ne "n")
	
	<#
	# Finalizing
	#>
	Write-Host "`n"
	
	if($global:pdfPath){
		Write-Host " PDF Report Location:    $global:pdfPath"
	}
	else{
		Write-Host " PDF Report Location:    No PDF File available"
	}
	if($htmlPath){
		Write-Host " HTML Report Location:   $htmlPath"
		Write-Host " 
    Automatically opening HTML Report now...
-----------------------------------------------------------------------------------"

		Invoke-Item "$htmlPath"
	}
	else{
		Write-Host " HTML Report Location:   No HTML File available"
	}

    if ($SkipDecline) {
		Write-Host "
    Actions:

    e) Return to main menu"

		do {
			$choice = Read-Host " Choose an Option"
			Write-Log " User Input: $choice"
			switch ($choice) {
				e { Show-Menu }
				default { 
					Write-Log " Wrong Input."
					Write-Host "Wrong Input. Please choose an option above." 
				}
			}

		} while ($choice -ne "e")
	}
	else{
		Write-Host "
    Actions:
    1) Delete Superseded Updates

    e) Return to main menu"
	
		do {
			$choice = Read-Host " Choose an Option"
			Write-Log " User Input: $choice"
			switch ($choice) {
				1 { Start-Run-DeleteSupersededUpdates }
				e { Show-Menu }
				default { 
					Write-Log " Wrong Input."
					Write-Host "Wrong Input. Please choose an option above." 
				}
			}

		} while ($choice -ne 1 -and $choice -ne "e")
	}
	
	# Finish logging	
	Write-Log "Finish 'WSUS Cleanup Report' Logging."

}

###
### Last Synchronization Report
###
function Start-Gen-LastWSUSSynchronizationReport {
	# Report Version
	$LastSyncReportVersion = "0.3"
	
	$CompName = $env:COMPUTERNAME
	$DateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
	
	# Start logging	
	Write-Log "Start 'Last Synchronization Report' Logging."
	
	# Create required directories
	<#Write-Log "Create required directories."
	$directories = @(
		"C:\_psc\psc_wsusreporting\Logfiles"
		"C:\_psc\psc_wsusreporting\Exports"
	)

	foreach ($dir in $directories) {
		Write-Log "Directory '$dir' already exists."
		If (-not (Test-Path $dir)) { 
			Write-Log "Creating Directory '$dir'."
			try{
				New-Item -Path $dir -ItemType Directory
			}
			catch{
				Write-Log "ERROR: Directory '$dir' could not be created."
			}
		}
	}
	#>
	
	$FileName = $Hostname+"_WU-Sync-Report_"+$DateTime+".html"
	$RootDir = "C:\_psc\psc_wsusreporting\"
	$FileDir = "C:\_psc\psc_wsusreporting\Exports\"
	$MediaDir = "C:\_psc\psc_wsusreporting\Exports\Media\"
	$FilePath = "$FileDir$FileName"

	$HTMLReportFileSelected = "false"
	$GetHTMLReportFiles = ""
	$SumOfHTMLReportFile = ""
	$HTMLReportFile = ""
	$HTMLReportFileList = @()

		
	
	function Start-WSUS-Reporting {

		# Function to extract Details of Update
		function Get-UpdateDetails($update, $approvalStatus) {
			if($update.IsSuperseded){
				$IsSuperseded = "Superseded"
			}
			else{
				$IsSuperseded = "-"
			}

			if($update.HasSupersededUpdates){
				$HasSuperseded = "Yes"
			}
			else{
				$HasSuperseded = "No"
			}
			
			$kb = ($update.KnowledgebaseArticles -join ", ")
			if($kb){
				$kb = "KB$($kb)"
			}
			else{
				$kb = "-"
			}
			$url = $update.AdditionalInformationUrls -join ", "
			$product = $update.ProductTitles -join ", "
            
            <#
			if($Approved){
				$UpdateAction = "OK"
			}
			else{
				$UpdateAction = "NOK"
			}
            #>

            # Use the passed approval status here
            if ($approvalStatus) {
                $UpdateAction = $approvalStatus
            }
            else {
                $UpdateAction = "-"
            }

			$PublishingDateFormatted = $update.CreationDate.ToString("dd MMM. yyyy - HH:mm:ss")
			$ArrivalDateFormatted = $update.ArrivalDate.ToString("dd MMM. yyyy - HH:mm:ss")

			return [PSCustomObject]@{
				ApprovedForInstallation = $UpdateAction
				PublishingDate = $PublishingDateFormatted #$update.CreationDate
				ArrivalDate = $ArrivalDateFormatted #$update.ArrivalDate
				ProductTitles = $product
				Classification = $update.UpdateClassificationTitle
				<#IsDeclined  = $update.IsDeclined
				PublicationState = $update.PublicationState
				HasSupersededUpdates = $HasSuperseded
				IsSuperseded = $IsSuperseded#>
				KBArticle   = $kb
				Title       = $update.Title
				SupportURL  = $url
			}
		}



		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 10
		Write-Log "Fetch WSUS Report Information."
		# Load WSUS Administration assembly
		try{
			[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
		}
		catch{
			Write-Log "ERROR: Can not fetch WSUS Information. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS Information. Reason: $_"
		}

		# Connect to local WSUS server
		Write-Log "Connect to local WSUS serve."
		try{
			$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
		}
		catch{
			Write-Log "ERROR: Can not connect to WSUS. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS. Reason: $_"
		}

		# Get last sync history entry
		# Get the synchronization history (use this, it's safe and works)
		Write-Log "Get sync history."
		try{
			$syncHistory = $wsus.GetSubscription().GetSynchronizationHistory() | Sort-Object StartTime -Descending
		}
		catch{
			Write-Log "ERROR: Can not get sync history. Reason: $_"
			Write-Warning "ERROR: Can not get sync history. Reason: $_"
		}

		# Get the most recent sync event
		Write-Log "Get the most recent sync event."
		try{
			$lastSync = $syncHistory | Select-Object -First 1
		}
		catch{
			Write-Log "ERROR: Can not get the most recent sync event. Reason: $_"
			Write-Warning "ERROR: Can not get the most recent sync event. Reason: $_"
		}
		
		if ($null -eq $lastSync -or [string]::IsNullOrWhiteSpace($lastSync.Result)) {
			$lastSyncResult = "No Data available"
		}
		else{
			$lastSyncResult = $lastSync.Result
		}
		
		# filter for a specific report
		#$lastSync = $syncHistory | Where-Object { $_.Id -eq "68bff289-d662-4e6d-aca6-9a3c48524714" }

		# Format the Start and End times
		Write-Log "Format the Start and End times."
		if($lastSync.StartTime -and $lastSync.EndTime){
			$startTimeFormatted = $lastSync.StartTime.ToString("dd MMM. yyyy - HH:mm")
			$endTimeFormatted = $lastSync.EndTime.ToString("dd MMM. yyyy - HH:mm")
		}
		else{
			$startTimeFormatted = "No Data available"
			$endTimeFormatted = "No Data available"
		}
		
		# Check result of last sync
		Write-Log "Check result of last sync."
		if ([string]::IsNullOrWhiteSpace($lastSync.Result)) {
			$lastSyncResult = "No Data available"
		}
		else{
			$lastSyncResult = $lastSync.Result
		}

		# Format Report creation time
		Write-Log "Format Report creation time."
		$ReportTime = Get-Date -Format "dd MMM. yyyy - HH:mm"

		# Get all updates downloaded/changed during the last sync
		Write-Log "Get all updates downloaded/changed during the last sync."
		$updates = $wsus.GetUpdates() | Where-Object {
			
			$_.ArrivalDate -ge $lastSync.StartTime 

			# If we filter by a specific snyc report id
			#$_.ArrivalDate -ge $lastSync.StartTime -and $_.ArrivalDate -le $lastSync.EndTime
		}

		# Separate new, revisioned and obsolete updates of the sync
		Write-Log "Separate new, revisioned and obsolete updates of the sync."
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 30

		$newUpdates = $updates | Where-Object { <#$_.IsDeclined -eq $false -and#> $_.PublicationState -ne "Expired" -and $_.HasEarlierRevision -eq $false -and $_.IsLatestRevision -eq $true}
		$obsoleteUpdates = $updates | Where-Object { $_.IsDeclined -eq $true -and $_.PublicationState -eq "Expired" }
		$revisionedUpdates = $updates | Where-Object { $_.PublicationState -ne "Expired" -and $_.IsDeclined -eq $false -and $_.HasEarlierRevision -eq $true -and $_.IsLatestRevision -eq $true }

    


		# Check Approved and Declined Updates since last sync
		Write-Log "Check Approved and Declined Updates since last sync."
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 40

		# Get the "All Computers" target group object once
		Write-Log "Get the 'All Computers' target group object once."
		$allComputersGroup = $wsus.GetComputerTargetGroups() | Select-Object Id,Name | Where-Object { $_.Name -eq "All Computers" -or $_.Name -eq "Alle Computer"}


		$approvedSinceLastSync = @()
		$declinedSinceLastSync = @()
		$approvals = @()
		$approvedUpdateIds = @{}
		$declinedUpdateIds = @{}
		$updateId = ""

		<#
		approval.Action can be:
			- Install
			- Uninstall
			- NotApproved
			- Decline
		#>
        
        <#
		foreach ($update in $updates) {
			$approvals = $update.GetUpdateApprovals()
			$updateGuid = $update.Id.UpdateId.Guid.ToString()

            # Track if this update has approved or declined already (unique)
            $hasApproved = $false
            $hasDeclined = $false
			
			foreach ($approval in $approvals) {
				# Only include approvals that were made after the last sync
				if ($approval.CreationDate -ge $lastSync.StartTime) {
                    #$approval.Action
					if ($approval.Action -eq "Install" -and -not $approvedUpdateIds.ContainsKey($updateGuid)) {
						$approvedUpdateIds[$updateGuid] = $true
						$approvedSinceLastSync += $update
						$Approved = $true
					}
					elseif ($approval.Action -eq "Decline" -and -not $declinedUpdateIds.ContainsKey($updateGuid)) {
						$declinedUpdateIds[$updateGuid] = $true
						$declinedSinceLastSync += $update
						$Approved = $false
					}
				}

                if ($approval.Action -eq "Install") {
                    $hasApproved = $true
                }
                elseif ($approval.Action -eq "Decline") {
                    $hasDeclined = $true
                }

			}
            
            if ($hasApproved -and -not $approvedUpdateIds.ContainsKey($updateGuid)) {
                $approvedUpdateIds[$updateGuid] = $true
                $approvedSinceLastSync += $update
            }

            if ($hasDeclined -and -not $declinedUpdateIds.ContainsKey($updateGuid)) {
                $declinedUpdateIds[$updateGuid] = $true
                $declinedSinceLastSync += $update
            }

		}#>

        # Step 1: Build approval status map for all updates
        $approvalStatusMap = @{}

        foreach ($update in $updates) {
            $updateGuid = $update.Id.UpdateId.Guid.ToString()

            # Initialize status to neutral
            $approvalStatusMap[$updateGuid] = "-"

            # Check if the update is declined directly
            <#if ($update.IsDeclined) {
                if (-not $declinedUpdateIds.ContainsKey($updateGuid)) {
                    $declinedUpdateIds[$updateGuid] = $true
                    $declinedSinceLastSync += $update
                    #$Approved = $false
                    #$approvalStatus = "NOK"
                    $approvalStatusMap[$updateGuid] = "NOK"
                }
            }
            else {
                # Not declined, so check approvals for Install action since last sync
                $approvalStatusMap[$updateGuid] = "Not Approved"  # default neutral, can change to "-"
                $approvals = $update.GetUpdateApprovals()
                foreach ($approval in $approvals) {
                    # Only consider approvals created since last sync started
                    if ($approval.CreationDate -ge $lastSync.StartTime) {
                        if ($approval.Action -eq "Install" -and -not $approvedUpdateIds.ContainsKey($updateGuid)) {
                            $approvedUpdateIds[$updateGuid] = $true
                            $approvedSinceLastSync += $update
                            #$Approved = $true
                            #$approvalStatus = "OK"
                            $approvalStatusMap[$updateGuid] = "OK"
                            #break # No need to check other approvals once approved
                        }
                    }
                }
            }#>

            $approvals = $update.GetUpdateApprovals()
            $wasApprovedBefore = $false
            $wasDeclined = $update.IsDeclined

            foreach ($approval in $approvals) {
                if ($approval.Action -eq "Install") {
                    $wasApprovedBefore = $true
                }
            }

            # Obsolete (expired) logic
            #Write-Host "PublicationState for $($updateGuid): '$($update.PublicationState)'"
            #Read-Host "Press any key"

            if ($update.PublicationState -eq "Expired") {
                if ($wasApprovedBefore -and $update.IsDeclined) {
                    $approvalStatusMap[$updateGuid] = "OBSOLETE-APPROVED-DECLINED"
                    #Read-Host "Expired, approved before, and now declined"
                }
                elseif ($wasApprovedBefore) {
                    $approvalStatusMap[$updateGuid] = "OBSOLETE-APPROVED"
                    #Read-Host "Expired and previously approved"
                }
                elseif ($update.IsDeclined -and $wasApprovedBefore -eq $false) {
                    $approvalStatusMap[$updateGuid] = "OBSOLETE"
                    #$approvalStatusMap[$updateGuid] = "OBSOLETE-DECLINED"
                    #Read-Host "Expired and declined"
                }
                else {
                    $approvalStatusMap[$updateGuid] = "OBSOLETE"
                    #Read-Host "Expired only"
                }

                continue
            }
            <#
            else {
                # Non-obsolete updates
                if ($update.IsDeclined) {
                    if (-not $declinedUpdateIds.ContainsKey($updateGuid)) {
                        $declinedUpdateIds[$updateGuid] = $true
                        $declinedSinceLastSync += $update
                        #$Approved = $false
                        #$approvalStatus = "NOK"
                        $approvalStatusMap[$updateGuid] = "NOK"
                    }
                    
                }
                else {
                    # Not declined, so check approvals for Install action since last sync
                    $approvalStatusMap[$updateGuid] = "-"  # default neutral, can change to "-"
                    $approvals = $update.GetUpdateApprovals()
                    foreach ($approval in $approvals) {
                        # Only consider approvals created since last sync started
                        if ($approval.CreationDate -ge $lastSync.StartTime) {
                            if ($approval.Action -eq "Install" -and -not $approvedUpdateIds.ContainsKey($updateGuid)) {
                                $approvedUpdateIds[$updateGuid] = $true
                                $approvedSinceLastSync += $update
                                #$Approved = $true
                                #$approvalStatus = "OK"
                                $approvalStatusMap[$updateGuid] = "OK"
                                #break # No need to check other approvals once approved
                            }
                        }
                    }
                }
            }
            #>

            # Handle currently declined (non-expired)
            if ($wasDeclined -and $update.PublicationState -ne "Expired") {
                if (-not $declinedUpdateIds.ContainsKey($updateGuid)) {
                    $declinedUpdateIds[$updateGuid] = $true
                    $declinedSinceLastSync += $update
                }
                $approvalStatusMap[$updateGuid] = "DECLINED"
                continue
            }

            # Handle newly approved since last sync
            foreach ($approval in $approvals) {
                if ($approval.CreationDate -ge $lastSync.StartTime -and $approval.Action -eq "Install") {
                    if (-not $approvedUpdateIds.ContainsKey($updateGuid)) {
                        $approvedUpdateIds[$updateGuid] = $true
                        $approvedSinceLastSync += $update
                    }
                    $approvalStatusMap[$updateGuid] = "OK"
                    break
                }
            }
        }

        #Test output
        <#
        # Output results for verification
        Write-Host "Approved Updates (since last sync): $($approvedSinceLastSync.Count)"
        Write-Host "Declined Updates (current): $($declinedSinceLastSync.Count)"

		
        Write-Host "Approved:"
		$approvedSinceLastSync | ft
        Write-Host "Declined:"
		$declinedSinceLastSync | ft
		
		Write-Host "
		+ New Updates:             $($newUpdates.Count)
		+ Revisioned Updates:      $($revisionedUpdates.Count)
		+ Obsolete Updates:        $($obsoleteUpdates.Count)

		+ Approved Updates (by Admin):   $($approvedSinceLastSync.Count)
		+ Declined Updates (by Admin):   $($declinedSinceLastSync.Count)
		"
		Write-Host "`n== Verification =="
		Write-Host "New Updates:        $($newUpdates.Count)"
		Write-Host "Approved (Unique):  $($approvedSinceLastSync.Count)"
		Write-Host "Declined (Unique):  $($declinedSinceLastSync.Count)"

		# Extra check: do any updates appear more than once?
		$dupeApproved = $approvedSinceLastSync | Group-Object -Property { $_.Id.UpdateId.Guid } | Where-Object { $_.Count -gt 1 }
		Write-Host "Duplicate approved updates found: $($dupeApproved.Count)"
		Read-Host "Press any key"
        #>

		# Remove duplicates, just to be sure
		Write-Log "Remove duplicates, just to be sure."
		#$approvedSinceLastSync = $approvedSinceLastSync | Sort-Object Id -Unique
		#$declinedSinceLastSync = $declinedSinceLastSync | Sort-Object Id -Unique
        
        # Now generate the output arrays, including the approval status
		#$approvedSinceLastSync = $approvedSinceLastSync | ForEach-Object { Get-UpdateDetails $_ "OK" }
		#$declinedSinceLastSync = $declinedSinceLastSync | ForEach-Object { Get-UpdateDetails $_ "NOK" }

		<#
        $newUpdates = $newUpdates | ForEach-Object { Get-UpdateDetails $_ "-" }
		$obsoleteUpdates = $obsoleteUpdates | ForEach-Object { Get-UpdateDetails $_ "-" }
		$revisionedUpdates = $revisionedUpdates | ForEach-Object { Get-UpdateDetails $_ "-" }
        
        $newUpdates = $newUpdates | ForEach-Object {
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        }

        $obsoleteUpdates = $obsoleteUpdates | ForEach-Object {
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        }

        $revisionedUpdates = $revisionedUpdates | ForEach-Object {
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        }
        #>
        $newUpdates = $newUpdates | ForEach-Object { 
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        } | Sort-Object -Property ApprovedForInstallation -Descending

        $obsoleteUpdates = $obsoleteUpdates | ForEach-Object { 
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        } | Sort-Object -Property ApprovedForInstallation -Descending

        $revisionedUpdates = $revisionedUpdates | ForEach-Object { 
            $guid = $_.Id.UpdateId.Guid.ToString()
            $status = if ($approvalStatusMap.ContainsKey($guid)) { $approvalStatusMap[$guid] } else { "-" }
            Get-UpdateDetails $_ $status
        } | Sort-Object -Property ApprovedForInstallation -Descending

		<#
		# Process and output
		#>
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 50

		Add-Content $FilePath -Value @"
    <h2>Report Information</h2>		
    <table class="report-info">
	<tbody>
	<tr>
	<td>Sync Start Time:</td>
	<td>$startTimeFormatted</td>
	</tr>
	<tr>
	<td>Sync End Time:</td>
	<td>$endTimeFormatted</td>
	</tr>
	<tr>
	<td>Report generated:</td>
	<td>$ReportTime</td>
	</tr>
    <tr>
    <td>WSUS Server:</td>
    <td>$CompName</td>
    </tr>
	<tr>
	<td>Result:</td>
	<td>$($lastSyncResult)</td>
	</tr>
	<tr>
	<td><a href="#new-updates">New Updates:</a></td>
	<td>$($newUpdates.Count)</td>
	</tr>
    <tr>
	<td><a href="#revisioned-updates">Revisioned Updates:</a></td>
	<td>$($revisionedUpdates.Count)</td>
	</tr>
    <tr>
	<td><a href="#obsolete-updates">Obsolete Updates:</a></td>
	<td>$($obsoleteUpdates.Count)</td>
	</tr>
    <tr>
	<td>Updates Approved for Installation:</td>
	<td>$($approvedSinceLastSync.Count)</td>
	</tr>
    <tr>
	<td>Declined Updates:</td>
	<td>$($declinedSinceLastSync.Count)</td>
	</tr>
    </tbody>
    </table>
"@

		#Create Table for New Updates
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 60
		Add-Content $FilePath -Value @"
    <h2 id="new-updates">New Updates Added</h2>
    <table class="list-syncedUpdates">
    <tbody>
    <tr>
    <th>Approved for Installation</th>
    <th>Publishing Date</th>
    <th>Arrival Date</th>
    <th>Product Titles</th>
    <th>Classification</th>
	<th>KBArticle</th>
	<th>Title</th>
	<th>SupportURL</th>
    </tr>
"@

		$UpdateApprovalList = @()
		$UpdatePublishingDateList = @() 
		$UpdateArrivalDateList = @()
		$UpdateProductTitleList = @()
		$UpdateClassificationList = @()
		$UpdateKBArticleList = @()
		$UpdateTileList = @()
		$UpdateSupportURLList = @()
    

		#Fill Table with content
		if($newUpdates){
			foreach ($NewUpdate in $newUpdates) {
				
				$UpdateApprovalList += $NewUpdate.ApprovedForInstallation
				$UpdatePublishingDateList += $NewUpdate.PublishingDate
				$UpdateArrivalDateList += $NewUpdate.ArrivalDate
				$UpdateProductTitleList += $NewUpdate.ProductTitles
				$UpdateClassificationList += $NewUpdate.Classification
				$UpdateKBArticleList += $NewUpdate.KBArticle
				$UpdateTileList += $NewUpdate.Title
				$UpdateSupportURLList += $NewUpdate.SupportURL

			}
			
			for(($x = 0); $x -lt $newUpdates.Count; $x++) {
				$ApprovalState = $UpdateApprovalList[$x]
				$PublishingDate = $UpdatePublishingDateList[$x]
				$ArrivalDate = $UpdateArrivalDateList[$x]
				$ProductTitle = $UpdateProductTitleList[$x]
				$Classification = $UpdateClassificationList[$x]
				$KBArticle = $UpdateKBArticleList[$x]
				$UpdateTitle = $UpdateTileList[$x]
				$SupportURL = $UpdateSupportURLList[$x]

				Add-Content $FilePath -Value @"
			<tr>
"@

				if($ApprovalState -eq "OK"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"> $ApprovalState</td>
"@
				}
				elseif($ApprovalState -eq "DECLINED"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/error.png" alt="Error Icon"> $ApprovalState</td>
"@
				}
                else{
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"> $ApprovalState</td>
"@
				}

				Add-Content $FilePath -Value @"
			<td>$PublishingDate</td>
			<td>$ArrivalDate</td>
			<td>$ProductTitle</td>
			<td>$Classification</td>
			<td>$KBArticle</td>
			<td>$UpdateTitle</td>
			<td class="Url"><a href="$SupportURL" target="_blank">$SupportURL</a></td>
			</tr>
"@
			}
		}
		else{
			Add-Content $FilePath -Value @"
			<tr>
			<td>No Data available</td>
			</tr>
"@
		}
		
		#Finish Table
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@

		#Create Table for Revisinoed Updates
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 70
		Add-Content $FilePath -Value @"
    <h2 id="revisioned-updates">Revisioned Updates</h2>
    <table class="list-syncedUpdates">
    <tbody>
    <tr>
    <th>Approved for Installation</th>
    <th>Publishing Date</th>
    <th>Arrival Date</th>
    <th>Product Titles</th>
    <th>Classification</th>
	<th>KBArticle</th>
	<th>Title</th>
	<th>SupportURL</th>
    </tr>
"@
    
		$UpdateApprovalList = @()
		$UpdatePublishingDateList = @() 
		$UpdateArrivalDateList = @()
		$UpdateProductTitleList = @()
		$UpdateClassificationList = @()
		$UpdateKBArticleList = @()
		$UpdateTileList = @()
		$UpdateSupportURLList = @()

		#Fill Table with content
		if($revisionedUpdates){
			foreach ($RevisionedUpdate in $revisionedUpdates) {
				
				$UpdateApprovalList += $RevisionedUpdate.ApprovedForInstallation
				$UpdatePublishingDateList += $RevisionedUpdate.PublishingDate
				$UpdateArrivalDateList += $RevisionedUpdate.ArrivalDate
				$UpdateProductTitleList += $RevisionedUpdate.ProductTitles
				$UpdateClassificationList += $RevisionedUpdate.Classification
				$UpdateKBArticleList += $RevisionedUpdate.KBArticle
				$UpdateTileList += $RevisionedUpdate.Title
				$UpdateSupportURLList += $RevisionedUpdate.SupportURL

			}
	
			for(($x = 0); $x -lt $revisionedUpdates.Count; $x++) {
				$ApprovalState = $UpdateApprovalList[$x]
				$PublishingDate = $UpdatePublishingDateList[$x]
				$ArrivalDate = $UpdateArrivalDateList[$x]
				$ProductTitle = $UpdateProductTitleList[$x]
				$Classification = $UpdateClassificationList[$x]
				$KBArticle = $UpdateKBArticleList[$x]
				$UpdateTitle = $UpdateTileList[$x]
				$SupportURL = $UpdateSupportURLList[$x]
				
				Add-Content $FilePath -Value @"
			<tr>
"@

				if($ApprovalState -eq "OK"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"> $ApprovalState</td>
"@
				}
				elseif($ApprovalState -eq "DECLINED"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/error.png" alt="Error Icon"> $ApprovalState</td>
"@
				}
                else{
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"> $ApprovalState</td>
"@
				}

				Add-Content $FilePath -Value @"
			<td>$PublishingDate</td>
			<td>$ArrivalDate</td>
			<td>$ProductTitle</td>
			<td>$Classification</td>
			<td>$KBArticle</td>
			<td>$UpdateTitle</td>
			<td class="Url"><a href="$SupportURL" target="_blank">$SupportURL</a></td>
			</tr>
"@
			}
		}
		else{
			Add-Content $FilePath -Value @"
			<tr>
			<td>No Data available</td>
			</tr>
"@
		}
		
		#Finish Table
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@

		#Create Table for Obsolete Updates
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 80
		Add-Content $FilePath -Value @"
    <h2 id="obsolete-updates">Obsolete and/or Expired Updates</h2>
    <table class="list-syncedUpdates">
    <tbody>
    <tr>
    <th>Approved for Installation</th>
    <th>Publishing Date</th>
    <th>Arrival Date</th>
    <th>Product Titles</th>
    <th>Classification</th>
	<th>KBArticle</th>
	<th>Title</th>
	<th>SupportURL</th>
    </tr>
"@

		$UpdateApprovalList = @()
		$UpdatePublishingDateList = @() 
		$UpdateArrivalDateList = @()
		$UpdateProductTitleList = @()
		$UpdateClassificationList = @()
		$UpdateKBArticleList = @()
		$UpdateTileList = @()
		$UpdateSupportURLList = @()

		#Fill Table with content
		if($obsoleteUpdates){
			foreach ($ObsoleteUpdate in $obsoleteUpdates) {
				
				$UpdateApprovalList += $ObsoleteUpdate.ApprovedForInstallation
				$UpdatePublishingDateList += $ObsoleteUpdate.PublishingDate
				$UpdateArrivalDateList += $ObsoleteUpdate.ArrivalDate
				$UpdateProductTitleList += $ObsoleteUpdate.ProductTitles
				$UpdateClassificationList += $ObsoleteUpdate.Classification
				$UpdateKBArticleList += $ObsoleteUpdate.KBArticle
				$UpdateTileList += $ObsoleteUpdate.Title
				$UpdateSupportURLList += $ObsoleteUpdate.SupportURL

			}
		
			for(($x = 0); $x -lt $obsoleteUpdates.Count; $x++) {
				$ApprovalState = $UpdateApprovalList[$x]
				$PublishingDate = $UpdatePublishingDateList[$x]
				$ArrivalDate = $UpdateArrivalDateList[$x]
				$ProductTitle = $UpdateProductTitleList[$x]
				$Classification = $UpdateClassificationList[$x]
				$KBArticle = $UpdateKBArticleList[$x]
				$UpdateTitle = $UpdateTileList[$x]
				$SupportURL = $UpdateSupportURLList[$x]
				
				Add-Content $FilePath -Value @"
			<tr>
"@

				if($ApprovalState -eq "OK"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"> $ApprovalState</td>
"@
				}
				elseif($ApprovalState -eq "DECLINED"){
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/error.png" alt="Error Icon"> $ApprovalState</td>
"@
				}
                else{
					Add-Content $FilePath -Value @"
			<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"> $ApprovalState</td>
"@
				}

				Add-Content $FilePath -Value @"
			<td>$PublishingDate</td>
			<td>$ArrivalDate</td>
			<td>$ProductTitle</td>
			<td>$Classification</td>
			<td>$KBArticle</td>
			<td>$UpdateTitle</td>
			<td class="Url"><a href="$SupportURL" target="_blank">$SupportURL</a></td>
			</tr>
"@
			}
		}
		else{
			Add-Content $FilePath -Value @"
			<tr>
			<td>No Data available</td>
			</tr>
"@
		}
		
		#Finish Table
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@
		
		### Print Output to Console ###
		Write-Log "Report Version                 $LastSyncReportVersion"
		Write-Log "Sync Start Time:               $startTimeFormatted"
		Write-Log "Sync End Time:                 $endTimeFormatted"
		Write-Log "Report generated:              $ReportTime"
		Write-Log "WSUS Server:                   $CompName"
		Write-Log "Result:                        $($lastSyncResult)"
		Write-Log "New Updates:                   $($newUpdates.Count)"
		Write-Log "Revisioned Updates:            $($revisionedUpdates.Count)"
		Write-Log "Obsolete Updates:              $($obsoleteUpdates.Count)"
		Write-Log "Approved Updates (by Admin):   $($approvedSinceLastSync.Count)"
		Write-Log "Declined Updates (by Admin):   $($declinedSinceLastSync.Count)"
    

		Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"

		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "              Report Information"
		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "
    + Report Version           $LastSyncReportVersion

    + Sync Start Time:         $startTimeFormatted
    + Sync End Time:           $endTimeFormatted
    + Report generated:        $ReportTime
    + WSUS Server:             $CompName
    + Result:                  $($lastSyncResult)
                        
    + New Updates:             $($newUpdates.Count)
    + Revisioned Updates:      $($revisionedUpdates.Count)
    + Obsolete Updates:        $($obsoleteUpdates.Count)

    + Approved Updates (by Admin):   $($approvedSinceLastSync.Count)
    + Declined Updates (by Admin):   $($declinedSinceLastSync.Count)
"

		#Write-Host "All updates:"
		#$updates | select-object * | Format-List
		#$updates | ForEach-Object { Get-UpdateDetails $_ } | Format-Table -AutoSize

		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Green "              New Updates Added:"
		#Write-Host "-----------------------------------------------------------------------------------"
		#$newupdates | select-object * | Format-List
		if($newUpdates){
			$newUpdates | Sort-Object -Property ApprovedForInstallation -Descending | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		}

		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Yellow "              Revisioned Updates:"
		#Write-Host "-----------------------------------------------------------------------------------"
		if($revisionedUpdates){
			$revisionedUpdates | Sort-Object -Property ApprovedForInstallation -Descending | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		}

		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Red "              Obsolete and/or Expired Updates:"
		#Write-Host "-----------------------------------------------------------------------------------"
		if($obsoleteUpdates){
			$obsoleteUpdates | Sort-Object -Property ApprovedForInstallation -Descending | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		}
		<#
		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Green "              Approved Updates for Installation:"
		#Write-Host "-----------------------------------------------------------------------------------"
		#$newupdates | select-object * | Format-List
		if($approvedSinceLastSync){
			$approvedSinceLastSync | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		} 
		
		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Red "              Declined Updates for Installation:"
		#Write-Host "-----------------------------------------------------------------------------------"
		#$newupdates | select-object * | Format-List
		if($declinedSinceLastSync){
			$declinedSinceLastSync | Format-Table -AutoSize 
		}
		else{
			Write-Host "None"
		}  
		#>
		
		
		
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Finalizing:" -PercentComplete 100
	}
	
	<#function SelectHTMLReportFile {	
		$HTMLReportFileList = @()
		$HTMLReportFileSelected = $false
		$Options = @()
		$Input = ""

		$GetHTMLReportFiles = Get-ChildItem -File "$($FileDir)" -Name -Include *.html
		$SumOfHTMLReportFile = $GetHTMLReportFiles.Count
		
		while($HTMLReportFileSelected -eq "false")
		{
			foreach($HTMLReportFile in $GetHTMLReportFiles) {
				#Write-Host $ReqFile
				$HTMLReportFileList += $HTMLReportFile
			}

			Write-Host "
 Select a Report File:"
			for(($x = 0); $x -lt $SumOfHTMLReportFile; $x++) {
				Write-Host " $($x))" $HTMLReportFileList[$x]
				$Options += $x
			}
			
			$Input = Read-Host " Please select an option (0/1/..)"
			Write-Log " User Input: $Input"
			if($Options -contains $Input){
				$HTMLReportFileSelected = "true"
				$GetFile = $HTMLReportFileList[$Input]
				$HTMLReportFileName = $GetFile #+".csv"
				Write-Host " Selected HTML Report File:" $HTMLReportFileName `n -ForegroundColor Yellow
				
			}
			elseif($Options -notcontains $Input) {
				Write-Host " Wrong Input" `n -ForegroundColor Red
				Write-Log " Wrong Input."
			}
		}
	}#>
	function SelectHTMLReportFile {	
		$Options = @()
		$Input = ""
		$HTMLReportFileList = @()  
		$HTMLReportFileSelected = $false

		$GetHTMLReportFiles = Get-ChildItem -File "$FileDir" -Name -Include *.html
		$SumOfHTMLReportFile = $GetHTMLReportFiles.Count

		while (-not $HTMLReportFileSelected) {
			foreach ($HTMLReportFile in $GetHTMLReportFiles) {
				$HTMLReportFileList += $HTMLReportFile
			}

			Write-Host "`n Select a Report File:"
			for (($x = 0); $x -lt $SumOfHTMLReportFile; $x++) {
				Write-Host "    $x) $($HTMLReportFileList[$x])"
				$Options += $x
			}

			$Input = Read-Host "`n Please select an option (0/1/..)"
			Write-Log " User Input: $Input"

			if ($Options -contains [int]$Input) {
				$HTMLReportFileSelected = $true
				$GetFile = $HTMLReportFileList[$Input]
				$global:HTMLReportFileName = $GetFile
				Write-Host "`n Selected HTML Report File: $global:HTMLReportFileName" -ForegroundColor Yellow
			}
			else {
				Write-Host "`n Wrong Input" -ForegroundColor Red
				Write-Log " Wrong Input."
			}
		}
	}

	function ConvertToPDF {

		#$reportname = Get-ChildItem -Path C:\_psc\WSUS Files\WSUS_Reporting\Exports\*.html -Name
		#$reportname = $global:HTMLReportFileName
		$global:pdfPath = "$FileDir$htmlfilename.pdf"
		
		$msedge = @(
			"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
			"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
		)
		$msedge = $msedge | Where-Object { Test-Path $_ } | Select-Object -First 1
		
		if (-not $msedge) {
			Write-Log "ERROR: Microsoft Edge not found. Please install Microsoft Edge."
			Write-Host -ForegroundColor Red "Microsoft Edge not found. Please install Microsoft Edge."
			return
		}
		else{
		
			Write-Log "Converting HTML Report to PDF"
			Write-Log "Selected File: $htmlPath"
			Write-Log "Converted File: $global:pdfPath"

			Start-Sleep 5
		
			$arguments = "--print-to-pdf=""$global:pdfPath"" --headless --page-size=A4 --disable-extensions --no-pdf-header-footer --disable-popup-blocking --run-all-compositor-stages-before-draw --disable-checker-imaging --landscape ""file:///$htmlPath"""
			Write-Log "Edge arguments: $arguments"
			
			try{
			
				#Start-Process "msedge.exe" -ArgumentList @("--headless","--print-to-pdf=""$global:pdfPath""", "--landscape", "--page-size=A4","--disable-extensions","--no-pdf-header-footer","--disable-popup-blocking","--run-all-compositor-stages-before-draw","--disable-checker-imaging", "file:///$htmlPath") -Wait
				Start-Process -FilePath "$msedge" -ArgumentList $arguments -Wait -NoNewWindow
			
			}
			catch{
				Write-Log "ERROR: Convering to pdf failed. Reason: $_"
				Write-Host -ForegroundColor yellow "ERROR: Convering to pdf failed. Reason: $_"
			}
			Start-Sleep 5
		}
	}
	
	function GenerateHTMLReport {
    
		#Wait 10 Seconds - System needs to start background services etc. after foregoing reboot.
		Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Generating Report:" -PercentComplete 0
		Start-Sleep -Seconds 10

		If (-Not ( Test-Path $RootDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'psc_wsusreporting'."
			New-Item -Path "C:\_psc\" -Name "psc_wsusreporting" -ItemType "directory"
		}

		If (-Not ( Test-Path $FileDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Exports'."
			New-Item -Path $RootDir -Name "Exports" -ItemType "directory"
		}

		If (-Not ( Test-Path $MediaDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Media'."
			New-Item -Path $FileDir -Name "Media" -ItemType "directory"
		}

		#Create HTML Report
		Write-Log "Create HTML Report."
		If (-Not ( Test-Path $FilePath ))
		{
			
			#Copy CSS Stylesheet and Images from DeploymentShare
			#Write-Log "Copy CSS Stylesheet and Images from DeploymentShare."
			#Copy-Item -Path "$global:FileShare\*" -Destination $global:MediaDir -Recurse -Force
			
			#Create File
			Write-Log "Create File."
			New-Item $FilePath -ItemType "file" | out-null
		
			#Add Content to the File
			Write-Log "Add Content to the File."
			Add-Content $FilePath -Value @"
<!doctype html>
<html>
	<head>
		<title>WSUS Synchronization Report for $Hostname</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="description" content="this is a Report of the synchronization of windows updates."/>
		<meta name="thumbnail" content=""/>
		<link rel="stylesheet" href="Media/styles.css" />
	</head>
	<body>
	<div id="main">
		<div id="title">
			<img id="default_logo" src="Media/Powershell_logo.png" alt="Logo">
			<h1 id="title">WSUS Synchronization Report for $Hostname</h1>
            <table class="report-info">
			<tbody>
			<tr>
			<td>Report Template Version:</td>
			<td>Version $LastSyncReportVersion</td>
			</tr>
			</tbody>
			</table>
        </div>
"@
		}

		###### Function Calls ######
		Start-WSUS-Reporting

		#Finish HTML Report
		Add-Content $FilePath -Value @"
</div>
</body>
<footer>
</footer>
</html>
"@

	}
	
	<#
	# Starting Script
	#>
	Clear-Host
	GenerateHTMLReport
	SelectHTMLReportFile
	# Get HTML Report information
	$htmlfilename = $global:HTMLReportFileName -split ".html" | select-object -First 1
	$htmlPath = "$FileDir$global:HTMLReportFileName"
		
	
	do {
		$choice = Read-Host "`n Do you want to convert a report to pdf? (y/n)"
		Write-Log " User Input: $choice"
		switch ($choice) {
			"y" { ConvertToPDF }
			"n" {  }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne "y" -and $choice -ne "n")
	
	<#
	# Finalizing
	#>
	#Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Finalizing:" -PercentComplete 98
	Write-Host "`n"
	
	if($global:pdfPath){
		Write-Host " PDF Report Location:    $global:pdfPath"
	}
	else{
		Write-Host " PDF Report Location:    No PDF File available"
	}
	if($htmlPath){
		Write-Host " HTML Report Location:   $htmlPath"
		Write-Host " 
    Automatically opening HTML Report now...
-----------------------------------------------------------------------------------"

		Invoke-Item "$htmlPath"
	}
	else{
		Write-Host " HTML Report Location:   No HTML File available"
	}
	
	
	Write-Host "
    Actions:

    e) Return to main menu"

    
	do {
		$choice = Read-Host " Choose an Option"
		Write-Log " User Input: $choice"
		switch ($choice) {
			#1 { GenerateHTMLReport }
			e { Show-Menu }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne 1 -and $choice -ne "e")
	
	# Finish logging	
	Write-Log "Finish 'Last Synchronization Report' Logging."
}

###
### Endpoint Status Report
###
function Start-Gen-EndpointStatusReport {
	# Report Version
	$EndpointStatusReportVersion = "0.2"
	
	$CompName = $env:COMPUTERNAME
	$DateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
	
	# Start logging	
	Write-Log "Start 'Endpoint Status Report' Logging."
	
	$FileName = $Hostname+"_EndpointStatus-Report_"+$DateTime+".html"
	$RootDir = "C:\_psc\psc_wsusreporting\"
	$FileDir = "C:\_psc\psc_wsusreporting\Exports\"
	$MediaDir = "C:\_psc\psc_wsusreporting\Exports\Media\"
	$FilePath = "$FileDir$FileName"

	$HTMLReportFileSelected = "false"
	$GetHTMLReportFiles = ""
	$SumOfHTMLReportFile = ""
	$HTMLReportFile = ""
	$HTMLReportFileList = @()
	
	function Start-Endpoint-Reporting {

		Write-Progress -id 1 -Activity "Generating Endpoint Status Report" -Status "Generating Report:" -PercentComplete 10
		Write-Log "Fetch WSUS Report Information."
		# Load WSUS Administration assembly
		try{
			[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
		}
		catch{
			Write-Log "ERROR: Can not fetch WSUS Information. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS Information. Reason: $_"
		}

		# Connect to local WSUS server
		Write-Log "Connect to local WSUS serve."
		try{
			$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
		}
		catch{
			Write-Log "ERROR: Can not connect to WSUS. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS. Reason: $_"
		}

		# Get all Endpoint Groups
		Write-Log "Get all Endpoint Groups."
		
		$EndpointGroupIDsList = @()
		$EndpointGroupNamesList = @()
		
		# Get all groups
		try{
			# Get all groups
			$groups = $wsus.GetComputerTargetGroups()
		}
		catch{
			Write-Log "ERROR: Can not fetch endpoint groups. Reason: $_"
			Write-Warning "ERROR: Can not fetch endpoint groups. Reason: $_"
		}

		
		$EndpointGroupIDsList = $groups | select-object Id
		$EndpointGroupNamesList = $groups | select-object Name
		

		# Format Report creation time
		Write-Log "Format Report creation time."
		$ReportTime = Get-Date -Format "dd MMM. yyyy - HH:mm"

		

		# Working through groups.
		Write-Log "Working through groups."
		Write-Progress -id 1 -Activity "Generating Endpoint Status Report" -Status "Generating Report:" -PercentComplete 30
		
		# Store summary for all groups
		$groupSummaries = @()
		$endpointSummaries = @()
		
		$groupIndex = 0
		$totalGroups = $groups.Count
		
		# Loop through each group except 'all computers'
		# The more computer objects are connected with wsus server, the more time consuming this process gets!
		Write-Log "Loop through each group except 'all computers'."
		Write-Log "The more computer objects are connected with wsus server, the more time consuming this process gets!"
		
		foreach ($group in $groups) {
			if ($group.Name -ne "Alle Computer" -and $group.Name -ne "All Computers") {
				$groupIndex++

				# Update progress bar
				Write-Progress -Activity "Processing WSUS Endpoint Groups. This may take a while..." `
								-Id 2 `
								-ParentId 1 `
								-Status "Working on group '$($group.Name)'" `
								-PercentComplete (($groupIndex / $totalGroups) * 100)
						
				# Scope to this group
				$scope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
				$scope.ComputerTargetGroups.Add($group) | Out-Null

				# Get computers in this group
				$comp = 0
				$computers = $wsus.GetComputerTargets($scope)
				$totalComputers = $computers.Count

				# Init counters
				$errorCount = 0
				$upToDate = 0
				$pending = 0
				$noStatus = 0

				foreach ($computer in $computers) {
					$comp++
					$ComputerName = $computer.FullDomainName
					
					# Update progress bar
					Write-Progress -Activity "Processing Endpoints in Group" `
									-Id 3 `
									-ParentId 2 `
									-Status "Working on '$ComputerName'" `
									-PercentComplete (($comp / $totalComputers) * 100)
					
					try {
						$summary = $computer.GetUpdateInstallationSummary()
						
						#$ComputerName = $computer.FullDomainName
						
						<#
						Available Objects:
							UnknownCount                
							NotApplicableCount          
							NotInstalledCount          
							DownloadedCount             
							InstalledCount              
							InstalledPendingRebootCount 
							FailedCount                 
						#>
						if (-not $summary) {
							$noStatus++
							$OverallUpdateStatus = "No Information"
							continue
						}

						if ($summary.FailedCount -gt 0) {
							$errorCount++
							$OverallUpdateStatus = "Updates with Errors"
						}
						elseif ($summary.NotInstalledCount -gt 0 -or $summary.DownloadedCount -gt 0 -or $summary.InstalledPendingRebootCount -gt 0) {
							$pending++
							$OverallUpdateStatus = "Pending Updates"
						}
						elseif ($summary.InstalledCount -gt 0 -and $summary.NotInstalledCount -eq 0 -and $summary.InstalledPendingRebootCount -eq 0) {
							$upToDate++
							$OverallUpdateStatus = "Up-to-Date"
						}
						else {
							$noStatus++
							$OverallUpdateStatus = "No Information"
						}
					} catch {
						$noStatus++
						$OverallUpdateStatus = "No Information"
					}
					<#Write-Host "  Name:                        $ComputerName"
					Write-Host "  Overall Update Status:       $OverallUpdateStatus"
					Write-Host "--------------------------------------------------"#>
					
					$endpointSummaries += [PSCustomObject]@{
						GroupName = $group.Name
						EndpointName = $ComputerName
						UpdateStatus = $OverallUpdateStatus
					}
				}
				<#Write-Host "  Members:             $($computers.Count)"
				Write-Host "  Errors:              $errorCount"
				Write-Host "  Pending:             $pending"
				Write-Host "  Up-to-date:          $upToDate"
				Write-Host "  No Status:           $noStatus"
				Write-Host "--------------------------------------------------"#>
		
				# Store the group summary in the array
				$groupSummaries += [PSCustomObject]@{
					GroupName   = $group.Name
					Total       = $computers.Count
					Errors      = $errorCount
					Pending     = $pending
					UpToDate    = $upToDate
					NoStatus    = $noStatus
				}
			}
		}

		<#
		# Process and output
		#>
		Write-Progress -id 1 -Activity "Generating Endpoint Status Report" -Status "Generating Report:" -PercentComplete 50
		Add-Content $FilePath -Value @"
    <h2>Report Information</h2>
	<div id="container">
    <table class="report-info">
	<tbody>
	<tr>
	<td>Report generated:</td>
	<td>$ReportTime</td>
	<td></td> <!-- Add this to balance the row -->
	</tr>
    <tr>
    <td>WSUS Server:</td>
    <td>$CompName</td>
	<td></td> <!-- Add this to balance the row -->
	</tr>
"@
		Write-Log "Report Version                 $EndpointStatusReportVersion"
		Write-Log "========== WSUS Group Summary =========="
		$grpID = 0
		foreach ($summary in $groupSummaries) {
			$grpID++
			Write-Log "Group:                                $($summary.GroupName)"
			Write-Log "  Total Endpoints:                    $($summary.Total)"
			Write-Log "  Endpoints with errors:              $($summary.Errors)"
			Write-Log "  Endpoints with pending updates:     $($summary.Pending)"
			Write-Log "  Endpoints Up-to-date:               $($summary.UpToDate)"
			Write-Log "  Endpoints without status:           $($summary.NoStatus)"
			Write-Log "---------------------------------------------"
			
			Add-Content $FilePath -Value @"
	<tr>
	<td><a href="#$($summary.GroupName)">$($summary.GroupName)</a></td>
	<td>
	<table class="nested">
		<tr>
		<td>Total Endpoints:</a></td>
		<td>$($summary.Total)</td>
		</tr>
		<tr>
		<td>Endpoints with errors:</a></td>
		<td>$($summary.Errors)</td>
		</tr>
		<tr>
		<td>Endpoints with pending updates:</td>
		<td>$($summary.Pending)</td>
		</tr>
		<tr>
		<td>Endpoints Up-to-date:</td>
		<td>$($summary.UpToDate)</td>
		</tr>
		<tr>
		<td>Endpoints without status:</td>
		<td>$($summary.NoStatus)</td>
		</tr>
	</table>
	</td>
	<td class="chart">
		<script type="text/javascript">
			google.charts.setOnLoadCallback(function() 
			{
				const errorCount = $($summary.Errors);
				const pendingCount = $($summary.Pending);
				const upToDateCount = $($summary.UpToDate);
				const noStatusCount = $($summary.NoStatus);
				
				const total = $($summary.Total);
				
				var data;
				var options;
				
				if(total === 0){
					// Dummy chart to show "No Data"
					data = google.visualization.arrayToDataTable([
						['Status', 'Count'],
						['No Data', 1]
					]);
					
					options = {
						legend: 'none',
						pieSliceText: 'label',
						backgroundColor: 'transparent',
						width: 300,
						height: 300,
						colors: [
							'#636363'   // No Status
						]
					};
				}
				else{
					data = google.visualization.arrayToDataTable([
					  ['Status', 'Count'],
					  ['Installed', upToDateCount],
					  ['Pending', pendingCount],
					  ['Error', errorCount],
					  ['No Status', noStatusCount]
					]);
					
					options = {
						//title: 'Update Status - $($_.Name)',
						legend: 'none',
						pieSliceText: 'value',
						backgroundColor: 'transparent',
						width: 300,
						height: 300,
						//pieHole: 0.4, //converts pie to donut
						colors: [
							'#23b010',  // Up-To-Date
							'#f2c522',  // Pending
							'#f22222',  // Error
							'#636363'   // No Status
						]
					};
				}
				
				

				var chart = new google.visualization.PieChart(document.getElementById('myPlot_$grpID'));
				chart.draw(data, options);
			});
		</script>
		<div id="myPlot_$grpID" class="pieChart"></div>
	</tr>
"@
		}
		
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
	</div>
"@

		#Create Table
		Write-Progress -id 1 -Activity "Generating Endpoint Status Report" -Status "Generating Report:" -PercentComplete 80
		
		foreach ($summary in $groupSummaries) {
			Add-Content $FilePath -Value @"
    <h2 id="$($summary.GroupName)">$($summary.GroupName)</h2>
    <table class="list-Updates">
    <tbody>
    <tr>
	<th>State</th>
    <th>Name of Endpoint</th>
    <th>Overall Update Status</th>
    </tr>
"@

			$EndpointNameList = @()
			$UpdateStatusList = @() 
		
			#Fill Table with content
			# Filter matching endpoints
			$filteredEndpoints = $endpointSummaries | Where-Object { $_.GroupName -eq $($summary.GroupName) } | select-object EndpointName,UpdateStatus | Sort-Object EndpointName
			
			# Fill the arrays
			if($filteredEndpoints){
				foreach ($endpoint in $filteredEndpoints) {
					$EndpointNameList += $endpoint.EndpointName
					$UpdateStatusList += $endpoint.UpdateStatus
				}
				
				for(($x = 0); $x -lt $EndpointNameList.Count; $x++) {
					$EndpointHostname = $EndpointNameList[$x]
					$EndpointUpdateStatus = $UpdateStatusList[$x]
					

					Add-Content $FilePath -Value @"
				<tr>
"@

					if($EndpointUpdateStatus -eq "Up-to-Date"){
						Add-Content $FilePath -Value @"
				<!--<td>&#9989 $EndpointUpdateStatus</td>-->
					<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"></td>
"@
					}
					elseif($EndpointUpdateStatus -eq "Pending Updates"){
						Add-Content $FilePath -Value @"
				<!--<td>&#8505 Pending Updates</td>-->
					<td><img class="icon" src="Media/Icons/warning.png" alt="Warn Icon"></td>
"@
					}
					elseif($EndpointUpdateStatus -eq "Updates with Errors"){
						Add-Content $FilePath -Value @"
				<!--<td>&#10060 Updates with Errors</td>-->
					<td><img class="icon" src="Media/Icons/error.png" alt="Error Icon"></td>
"@
					}
					else{
						Add-Content $FilePath -Value @"
				<!--<td>&#9888 No Information</td>-->
					<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"></td>
"@
					}

					Add-Content $FilePath -Value @"
				<td>$EndpointHostname</td>
				<td>$EndpointUpdateStatus</td>
				</tr>
"@
				}
			}
			else{
				Add-Content $FilePath -Value @"
				<tr>
				<td>No Data available</td>
				</tr>
"@
			}
		
			#Finish Table
			Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@
		}


		Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"

		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "              Report Information"
		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "
    + Report Version           $EndpointStatusReportVersion
"
		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Yellow "              WSUS Group Summary"

		foreach ($summary in $groupSummaries) {
			Write-Host "
    + Group                                   $($summary.GroupName)
        Total endpoints:                      $($summary.Total)
        Endpoints with errors:                $($summary.Errors)
        Endpoints with pending updates:       $($summary.Pending)
        Endpoints Up-to-date:                 $($summary.UpToDate)
        Endpoints without status:             $($summary.NoStatus)
"
			<#Write-Host "`nGroup: $($summary.GroupName)" -ForegroundColor Cyan
			Write-Host "  Total Computers:     $($summary.Total)"
			Write-Host "  Errors:              $($summary.Errors)"
			Write-Host "  Pending Updates:     $($summary.Pending)"
			Write-Host "  Up-to-date:          $($summary.UpToDate)"
			Write-Host "  No Status:           $($summary.NoStatus)"
			Write-Host "---------------------------------------------"#>
		}
		
		foreach ($summary in $groupSummaries) {
			Write-Host "`n-----------------------------------------------------------------------------------"
			Write-Host -ForegroundColor Cyan "              Endpoint-Group $($summary.GroupName):"
			#Write-Host "-----------------------------------------------------------------------------------"
			if($endpointSummaries){
				$endpointSummaries | Where-Object { $_.GroupName -eq $($summary.GroupName) } | select-object EndpointName,UpdateStatus | Sort-Object EndpointName | Format-Table -AutoSize 
			}
			else{
				Write-Host "None"
			}
		}
		
		
		Write-Progress -id 1 -Activity "Generating Endpoint Status Report" -Status "Finalizing:" -PercentComplete 100
	}
	
	function SelectHTMLReportFile {	
		$Options = @()
		$Input = ""
		$HTMLReportFileList = @()  
		$HTMLReportFileSelected = $false

		$GetHTMLReportFiles = Get-ChildItem -File "$FileDir" -Name -Include *.html
		$SumOfHTMLReportFile = $GetHTMLReportFiles.Count

		while (-not $HTMLReportFileSelected) {
			foreach ($HTMLReportFile in $GetHTMLReportFiles) {
				$HTMLReportFileList += $HTMLReportFile
			}

			Write-Host "`n Select a Report File:"
			for (($x = 0); $x -lt $SumOfHTMLReportFile; $x++) {
				Write-Host "    $x) $($HTMLReportFileList[$x])"
				$Options += $x
			}

			$Input = Read-Host "`n Please select an option (0/1/..)"
			Write-Log " User Input: $Input"

			if ($Options -contains [int]$Input) {
				$HTMLReportFileSelected = $true
				$GetFile = $HTMLReportFileList[$Input]
				$global:HTMLReportFileName = $GetFile
				Write-Host "`n Selected HTML Report File: $global:HTMLReportFileName" -ForegroundColor Yellow
			}
			else {
				Write-Host "`n Wrong Input" -ForegroundColor Red
				Write-Log " Wrong Input."
			}
		}
	}

	function ConvertToPDF {

		#$reportname = Get-ChildItem -Path C:\_psc\WSUS Files\WSUS_Reporting\Exports\*.html -Name
		#$reportname = $global:HTMLReportFileName
		$global:pdfPath = "$FileDir$htmlfilename.pdf"
		
		$msedge = @(
			"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
			"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
		)
		$msedge = $msedge | Where-Object { Test-Path $_ } | Select-Object -First 1
		
		if (-not $msedge) {
			Write-Log "ERROR: Microsoft Edge not found. Please install Microsoft Edge."
			Write-Host -ForegroundColor Red "Microsoft Edge not found. Please install Microsoft Edge."
			return
		}
		else{
		
			Write-Log "Converting HTML Report to PDF"
			Write-Log "Selected File: $htmlPath"
			Write-Log "Converted File: $global:pdfPath"

			Start-Sleep 5
		
			$arguments = "--print-to-pdf=""$global:pdfPath"" --headless --page-size=A4 --disable-extensions --no-pdf-header-footer --disable-popup-blocking --run-all-compositor-stages-before-draw --disable-checker-imaging --landscape ""file:///$htmlPath"""
			Write-Log "Edge arguments: $arguments"
			
			try{
			
				#Start-Process "msedge.exe" -ArgumentList @("--headless","--print-to-pdf=""$global:pdfPath""", "--landscape", "--page-size=A4","--disable-extensions","--no-pdf-header-footer","--disable-popup-blocking","--run-all-compositor-stages-before-draw","--disable-checker-imaging", "file:///$htmlPath") -Wait
				Start-Process -FilePath "$msedge" -ArgumentList $arguments -Wait -NoNewWindow
			
			}
			catch{
				Write-Log "ERROR: Convering to pdf failed. Reason: $_"
				Write-Host -ForegroundColor yellow "ERROR: Convering to pdf failed. Reason: $_"
			}
			Start-Sleep 5
		}
	}
	
	function GenerateHTMLReport {
    
		#Wait 10 Seconds - System needs to start background services etc. after foregoing reboot.
		Write-Progress -id 1 -Activity "Generating WSUS Endpoint Report" -Status "Generating Report:" -PercentComplete 0
		Start-Sleep -Seconds 10

		If (-Not ( Test-Path $RootDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'psc_wsusreporting'."
			New-Item -Path "C:\_psc\" -Name "psc_wsusreporting" -ItemType "directory"
		}

		If (-Not ( Test-Path $FileDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Exports'."
			New-Item -Path $RootDir -Name "Exports" -ItemType "directory"
		}

		If (-Not ( Test-Path $MediaDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Media'."
			New-Item -Path $FileDir -Name "Media" -ItemType "directory"
		}

		#Create HTML Report
		Write-Log "Create HTML Report."
		If (-Not ( Test-Path $FilePath ))
		{
			
			#Copy CSS Stylesheet and Images from DeploymentShare
			#Write-Log "Copy CSS Stylesheet and Images from DeploymentShare."
			#Copy-Item -Path "$global:FileShare\*" -Destination $global:MediaDir -Recurse -Force
			
			#Create File
			Write-Log "Create File."
			New-Item $FilePath -ItemType "file" | out-null
		
			#Add Content to the File
			Write-Log "Add Content to the File."
			Add-Content $FilePath -Value @"
<!doctype html>
<html>
	<head>
		<title>WSUS Endpoint Report for $Hostname</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="description" content="this is a Report of endpoints connected with windows update server."/>
		<meta name="thumbnail" content=""/>
		<link rel="stylesheet" href="Media/styles.css" />
		<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
		<script>google.charts.load('current', {'packages':['corechart']});</script>
	</head>
	<body>
	<div id="main">
		<div id="title">
			<img id="default_logo" src="Media/Powershell_logo.png" alt="Logo">
			<h1 id="title">WSUS Endpoint Report for $Hostname</h1>
            <table class="report-info">
			<tbody>
			<tr>
			<td>Report Template Version:</td>
			<td>Version $EndpointStatusReportVersion</td>
			</tr>
			</tbody>
			</table>
        </div>
"@
		}

		###### Function Calls ######
		Start-Endpoint-Reporting

		#Finish HTML Report
		Add-Content $FilePath -Value @"
</div>
</body>
<footer>
</footer>
</html>
"@

	}
	
	<#
	# Starting Script
	#>
	Clear-Host
	GenerateHTMLReport
	SelectHTMLReportFile
	# Get HTML Report information
	$htmlfilename = $global:HTMLReportFileName -split ".html" | select-object -First 1
	$htmlPath = "$FileDir$global:HTMLReportFileName"
		
	
	do {
		$choice = Read-Host "`n Do you want to convert a report to pdf? (y/n)"
		Write-Log " User Input: $choice"
		switch ($choice) {
			"y" { ConvertToPDF }
			"n" {  }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne "y" -and $choice -ne "n")
	
	<#
	# Finalizing
	#>
	#Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Finalizing:" -PercentComplete 98
	Write-Host "`n"
	
	if($global:pdfPath){
		Write-Host " PDF Report Location:    $global:pdfPath"
	}
	else{
		Write-Host " PDF Report Location:    No PDF File available"
	}
	if($htmlPath){
		Write-Host " HTML Report Location:   $htmlPath"
		Write-Host " 
    Automatically opening HTML Report now...
-----------------------------------------------------------------------------------"

		Invoke-Item "$htmlPath"
	}
	else{
		Write-Host " HTML Report Location:   No HTML File available"
	}
	
	
	Write-Host "
    Actions:
	
    e) Return to main menu"

    
	do {
		$choice = Read-Host " Choose an Option"
		Write-Log " User Input: $choice"
		switch ($choice) {
			#1 { GenerateHTMLReport }
			e { Show-Menu }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne 1 -and $choice -ne "e")
	
	# Finish logging	
	Write-Log "Finish 'Endpoint Status Report' Logging."
}

###
### Update Status Report
###
function Start-Gen-UpdateStatusReport {
	# Report Version
	$UpdateStatusReportVersion = "0.1"
	
	$ReportName = "UpdateStatus-Report"
	
	$CompName = $env:COMPUTERNAME
	$DateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
	
	# Start logging	
	Write-Log "Start 'Update Status Report' Logging."
	
	$FileName = $Hostname+"_"+$ReportName+"_"+$DateTime+".html"
	$RootDir = "C:\_psc\psc_wsusreporting\"
	$FileDir = "C:\_psc\psc_wsusreporting\Exports\"
	$MediaDir = "C:\_psc\psc_wsusreporting\Exports\Media\"
	$FilePath = "$FileDir$FileName"

	$HTMLReportFileSelected = "false"
	$GetHTMLReportFiles = ""
	$SumOfHTMLReportFile = ""
	$HTMLReportFile = ""
	$HTMLReportFileList = @()
	
	function Start-Update-Reporting {

		Write-Progress -id 1 -Activity "Generating Updates Status Report" -Status "Generating Report:" -PercentComplete 10
		Write-Log "Fetch WSUS Report Information."
		# Load WSUS Administration assembly
		try{
			[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
		}
		catch{
			Write-Log "ERROR: Can not fetch WSUS Information. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS Information. Reason: $_"
		}

		# Connect to local WSUS server
		Write-Log "Connect to local WSUS serve."
		try{
			$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
		}
		catch{
			Write-Log "ERROR: Can not connect to WSUS. Reason: $_"
			Write-Warning "ERROR: Can not fetch WSUS. Reason: $_"
		}
		
		# Get all updates available on the WSUS server
		try {
			$updates = $wsus.GetUpdates()
		}
		catch {
			Write-Log "ERROR: Can not fetch updates. Reason: $_"
			Write-Warning "ERROR: Can not fetch updates. Reason: $_"
		}

		# Get all update classifications
		try {
			$classifications = $wsus.GetUpdateClassifications()
		}
		catch {
			Write-Log "ERROR: Can not fetch classifications. Reason: $_"
			Write-Warning "ERROR: Can not fetch classifications. Reason: $_"
		}

		# Get all computer target groups
		try {
			$groups = $wsus.GetComputerTargetGroups()
		}
		catch {
			Write-Log "ERROR: Can not fetch endpoint groups. Reason: $_"
			Write-Warning "ERROR: Can not fetch endpoint groups. Reason: $_"
		}
		
		# Get all computer targets # should be obsolete
		<#try {
			$computers = $wsus.GetComputerTargets()
		}
		catch {
			Write-Warning "ERROR: Can not fetch endpoints. Reason: $_"
		}#>

		# Create dictionary of approved updates created in the last x days
		$updatesDict = @{}
		$gotUpds = $false
		do {
			# Prompt the user for the number of days to filter updates
			$days = Read-Host "`n Enter the number of days that the report should consider (e.g., 30 for the last 30 days)
 Note: The larger your timespan, the longer it takes to generae the report!"

			# Validate the input
			if (-not ($days -match '^\d+$' -and $days -gt 0)) {
				Write-Host -ForegroundColor Red " Invalid input. Please enter a positive number."
				$gotUpds = $false
			}
			else {
				$cutoffDate = (Get-Date).AddDays(-$days)

				# Initialize the dictionary again to ensure it's cleared each time the loop runs
				$updatesDict.Clear()

				# Loop through all updates and filter based on approval status and creation date
				foreach ($update in $updates) {
					if ($update.IsApproved -eq $true -and $update.CreationDate -gt $cutoffDate) {
						$updclassification = ($update.UpdateClassificationTitle -join ', ')
						$updatesDict[$update.Id.UpdateId] = [PSCustomObject]@{
							Title          = $update.Title
							Classification = ($update.UpdateClassificationTitle -join ', ')
							#SupportURL = ($update.AdditionalInformationUrls -join ', ')
						}
					}
				}

				# Check if updates were found
				<#if ($updatesDict.Count -eq 0) {
					Write-Host -ForegroundColor Yellow " There are no updates to report within your specified timespan."
					$gotUpds = $false
				}
				else {
					$gotUpds = $true
				}#>
				$gotUpds = $true
			}
		}
		while (-not $gotUpds)  # Keep looping until valid input and updates are found
		$totalApprUpdates = $updatesDict.Count

		$groupIndex = 0
		$totalGroups = $groups.Count

		$classificationSummaries = @()
		$groupSummaries = @()
		$updSummaries = @()
		
		# Format Report creation time
		Write-Log "Format Report creation time."
		$ReportTime = Get-Date -Format "dd MMM. yyyy - HH:mm"
		
		# Iterate through each group and gather information
		foreach ($group in $groups) {
			if ($group.Name -ne "Alle Computer" -and $group.Name -ne "All Computers") {
			#if ($group.Name -eq "Productive") {
				$groupIndex++

				# Update progress bar
				Write-Progress -Activity "Processing WSUS Updates Summary. This may take a while..." `
								-Id 1 `
								-Status "Working on group '$($group.Name)'" `
								-PercentComplete (($groupIndex / $totalGroups) * 100)

				$groupSummary = @{
					GroupName      = $group.Name
					Classifications = @()
				}

				# Scope to this group
				$scope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
				$scope.ComputerTargetGroups.Add($group) | Out-Null

				# Get computers in this group
				$computers = $wsus.GetComputerTargets($scope)
				$totalComputers = $computers.Count
				$compIndex = 0

				$totalClasses = $classifications.Count
				$classIndex = 0

				# Iterate through each classification
				foreach ($class in $classifications) {
					$classIndex++

					# Update progress bar
					Write-Progress -Activity "Processing Classifications" `
									-Id 2 `
									-ParentId 1 `
									-Status "Working on classification '$($class.Title)'" `
									-PercentComplete (($classIndex / $totalClasses) * 100)

					# Scope to this class
					$scope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
					$scope.Classifications.Add($class) | out-null
		  

					# Initialize counters for update statuses
					$installedCount = 0
					$notApplicableCount = 0
					$downloadedCount = 0
					$notInstalled = 0
					$pendingCount = 0
					$failedCount = 0
					$noStatusCount = 0

					# Get Updates in this class
					$upds = $wsus.GetUpdates($scope)
					# Filter updates to only include approved updates in the selected date range and that match the current classification
					$filteredUpdates = $upds | Where-Object {
						$_.IsApproved -and $updatesDict.ContainsKey($_.Id.UpdateId)
					}


					$updIndex = 0
					#$totalUpds = $upds.Count
					$totalUpds = $filteredUpdates.Count

					# Loop through each update
					#foreach ($update in $upds) {
					foreach ($update in $filteredUpdates) {

						#if ($update.IsApproved -eq $true) { # Obsolete check
							$updIndex++

							# Update progress bar
							Write-Progress -Activity "Processing Updates" `
											-Id 3 `
											-ParentId 2 `
											-Status "Working on update '$($update.Title)'" `
											-PercentComplete (($updIndex / $totalUpds) * 100)

							$updateTitle = $update.Title
							

							try {
								# Get installation information for each update
								$installationInfo = $update.GetUpdateInstallationInfoPerComputerTarget($group)

								if ($null -ne $installationInfo.UpdateInstallationState) {
									switch ($installationInfo.UpdateInstallationState) {
										'Installed'       { $updateStatus = 'Installed'; $updateStatusID = 0 }
										'NotApplicable'   { $updateStatus = 'Not Needed'; $updateStatusID = 1 }
										'Downloaded'      { $updateStatus = 'Downloaded'; $updateStatusID = 2 }
										'NotInstalled'    { $updateStatus = 'Needed (But not installed yet)'; $updateStatusID = 3 }
										'Pending'         { $updateStatus = 'Pending (Not downloaded or installed yet)'; $updateStatusID = 4 }
										'Failed'          { $updateStatus = 'Failed'; $updateStatusID = 5 }
										'Unknown'         { $updateStatus = 'No Status Reported'; $updateStatusID = 6 }
										default           { $updateStatus = 'No Status Reported'; $updateStatusID = 6 }
									}
								} else {
									$updateStatus = 'No Information'
								}

								# Update counters based on the status
								switch ($updateStatusID) {
									0 { $installedCount++ }
									1 { $notApplicableCount++ }
									2 { $downloadedCount++ }
									3 { $notInstalled++ }
									4 { $pendingCount++ }
									5 { $failedCount++ }
									6 { $noStatusCount++ }
								}
							}
							catch {
								$noStatusCount++
								$updateStatus = "No Information"
							}

							$updSummaries += [PSCustomObject]@{
								ClassName = $class.Title
								UpdateName = $updateTitle
								UpdateStatus = $updateStatus
								UpdateStatusID = $updateStatusID
								SupportURL = ($update.AdditionalInformationUrls -join ', ')
							}
						#}
					}

					# Add classification summary to the group summary
					$groupSummary.Classifications += [PSCustomObject]@{
						Classification = $class.Title
						TotalUpdates   = $totalUpds
						Installed      = $installedCount
						NotApplicable  = $notApplicableCount
						Downloaded     = $downloadedCount
						NotInstalled   = $notInstalled
						Pending        = $pendingCount
						Failed         = $failedCount
						NoStatus       = $noStatusCount
					}

					# Also add to global classification summary list for report output
					$classificationSummaries += [PSCustomObject]@{
						ClassName       = $class.Title
						TotalUpdates    = $totalUpds
						Installed       = $installedCount
						NotApplicable   = $notApplicableCount
						Downloaded      = $downloadedCount
						NotInstalled    = $notInstalled
						Pending         = $pendingCount
						Failed          = $failedCount
						NoStatus        = $noStatusCount
					}
				}

				# Add the group summary to the overall summary
				$groupSummaries += $groupSummary
			}
		}



		<#
		# Process and output
		#>
		Write-Progress -id 1 -Activity "Generating Updates Status Report" -Status "Generating Report:" -PercentComplete 50
		Add-Content $FilePath -Value @"
    <h2>Report Information</h2>
	<div id="container">
    <table class="report-info">
	<tbody>
	<tr>
	<td>Report generated:</td>
	<td>$ReportTime</td>
	<td></td> <!-- Add this to balance the row -->
	</tr>
	<tr>
	<td>Report Scope:</td>
	<td>$days Days</td>
	<td></td> <!-- Add this to balance the row -->
	</tr>
    <tr>
    <td>WSUS Server:</td>
    <td>$CompName</td>
	<td></td> <!-- Add this to balance the row -->
	</tr>
"@
		Write-Log "Report Version                 $UpdateStatusReportVersion"
		Write-Log "========== WSUS Updates Summary =========="
		$classID = 0
		$classificationSummaries |
			Group-Object ClassName |
			ForEach-Object {
				$classID++
				$classGroup = $_.Group

				$totals = @{
					TotalUpdates     = ($classGroup | Measure-Object -Property TotalUpdates -Sum).Sum
					Installed        = ($classGroup | Measure-Object -Property Installed -Sum).Sum
					NotApplicable    = ($classGroup | Measure-Object -Property NotApplicable -Sum).Sum
					Downloaded       = ($classGroup | Measure-Object -Property Downloaded -Sum).Sum
					NotInstalled     = ($classGroup | Measure-Object -Property NotInstalled -Sum).Sum
					Pending          = ($classGroup | Measure-Object -Property Pending -Sum).Sum
					Failed           = ($classGroup | Measure-Object -Property Failed -Sum).Sum
					NoStatus         = ($classGroup | Measure-Object -Property NoStatus -Sum).Sum
				}

				Write-Log "Classification    $($_.Name)"
				Write-Log "    Total Updates:     $($totals.TotalUpdates)"
				Write-Log "    Installed:         $($totals.Installed)"
				Write-Log "    Not Applicable:    $($totals.NotApplicable)"
				Write-Log "    Downloaded:        $($totals.Downloaded)"
				Write-Log "    Not Installed:     $($totals.NotInstalled)"
				Write-Log "    Pending:           $($totals.Pending)"
				Write-Log "    Failed:            $($totals.Failed)"
				Write-Log "    No Status:         $($totals.NoStatus)"
			
			
				Add-Content $FilePath -Value @"
	<tr>
	<td><a href="#$($_.Name)">$($_.Name)</a></td>
	<td>
	<table class="nested">
		<tr>
		<td>Total Updates:</a></td>
		<td>$($totals.TotalUpdates)</td>
		</tr>
		<tr>
		<td>Updates Installed:</a></td>
		<td>$($totals.Installed)</td>
		</tr>
		<tr>
		<td>Updates Not Applicable:</td>
		<td>$($totals.NotApplicable)</td>
		</tr>
		<tr>
		<td>Updates Downloaded:</td>
		<td>$($totals.Downloaded)</td>
		</tr>
		<tr>
		<td>Updates Not Installed:</td>
		<td>$($totals.NotInstalled)</td>
		</tr>
		<tr>
		<td>Updates Pending Installation:</td>
		<td>$($totals.Pending)</td>
		</tr>
		<tr>
		<td>Updates Failed Installation:</td>
		<td>$($totals.Failed)</td>
		</tr>
		<tr>
		<td>Updates without Status:</td>
		<td>$($totals.NoStatus)</td>
		</tr>
	</table>
	</td>
	<td class="chart">
		<script type="text/javascript">
			google.charts.setOnLoadCallback(function() 
			{
				const installedCount = $($totals.Installed) + $($totals.NotApplicable);
				const pendingCount = $($totals.Downloaded) + $($totals.NotInstalled) + $($totals.Pending);
				const errorCount = $($totals.Failed);
				const noStatCount = $($totals.NoStatus);
				
				const total = installedCount + pendingCount + errorCount + noStatCount;
				
				var data;
				var options;
				
				if(total === 0){
					// Dummy chart to show "No Data"
					data = google.visualization.arrayToDataTable([
						['Status', 'Count'],
						['No Data', 1]
					]);
					
					options = {
						legend: 'none',
						pieSliceText: 'label',
						backgroundColor: 'transparent',
						width: 300,
						height: 300,
						colors: [
							'#636363'   // No Status
						]
					};
				}
				else{
					data = google.visualization.arrayToDataTable([
					  ['Status', 'Count'],
					  ['Installed', installedCount],
					  ['Pending', pendingCount],
					  ['Failed', errorCount],
					  ['No Status', noStatCount]
					]);
					
					options = {
						//title: 'Update Status - $($_.Name)',
						legend: 'none',
						pieSliceText: 'value',
						backgroundColor: 'transparent',
						width: 300,
						height: 300,
						//pieHole: 0.4, //converts pie to donut
						colors: [
							'#23b010',  // Installed
							'#f2c522',  // Pending
							'#f22222',  // Failed
							'#636363'   // No Status
						]
					};
				}
				
				

				var chart = new google.visualization.PieChart(document.getElementById('myPlot_$classID'));
				chart.draw(data, options);
			});
		</script>
		<div id="myPlot_$classID" class="pieChart"></div>
	</td>
	</tr>
"@
			}
		
		Add-Content $FilePath -Value @"
    </tbody>
    </table>
	</div>
"@

		#Create Table
		Write-Progress -id 1 -Activity "Generating Updates Status Report" -Status "Generating Report:" -PercentComplete 80
		
		$classificationSummaries |
			Group-Object ClassName |
			ForEach-Object {
				$className = $_.Name  # Save outer loop variable
				Add-Content $FilePath -Value @"
    <h2 id="$className">$className</h2>
    <table class="list-Updates">
    <tbody>
    <tr>
	<th>State</th>
	<th>Overall Update Status</th>
    <th>Update Title</th>
    <th>Support URL</th>
    </tr>
"@
				$UpdateNameList = @()
				$UpdateStatusList = @()
				$UpdateStatusIDList = @()
				$UpdateURLList = @()
				
				# Fill the arrays
				if($updSummaries){
					$filteredUpdates = $updSummaries | Where-Object { $_.ClassName -eq $className } | Select-Object UpdateStatusID,UpdateStatus,UpdateName,SupportURL | Sort-Object UpdateName
					
					if ($filteredUpdates.Count -gt 0) {
						foreach ($updateItem in $filteredUpdates) {
							$UpdateNameList += $updateItem.UpdateName
							$UpdateStatusList += $updateItem.UpdateStatus
							$UpdateStatusIDList += $updateItem.UpdateStatusID
							$UpdateURLList += $updateItem.SupportURL
						}
						
						for(($x = 0); $x -lt $UpdateNameList.Count; $x++) {
							$UpdateName = $UpdateNameList[$x]
							$UpdateInstallStatus = $UpdateStatusList[$x]
							$UpdSupportURL = $UpdateURLList[$x]
							$StatusID = $UpdateStatusIDList[$x]

							Add-Content $FilePath -Value @"
						<tr>
"@

							if($StatusID -eq 0){
								Add-Content $FilePath -Value @"
						<!--<td>&#9989 Installed</td>-->
							<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"></td>
							<td>Installed</td>
"@
							}
							elseif($StatusID -eq 1){
								Add-Content $FilePath -Value @"
						<!--<td>&#9989 Not Needed</td>-->
							<td><img class="icon" src="Media/Icons/check.png" alt="Check Icon"></td>
							<td>Not Needed</td>
"@
							}
							elseif($StatusID -eq 2){
								Add-Content $FilePath -Value @"
						<!--<td>&#8505 Downloaded</td>-->
							<td><img class="icon" src="Media/Icons/warning.png" alt="Warn Icon"></td>
							<td>Downloaded</td>
"@
							}
							elseif($StatusID -eq 3){
								Add-Content $FilePath -Value @"
						<!--<td>&#8505 Needd (But not installed yet)</td>-->
							<td><img class="icon" src="Media/Icons/warning.png" alt="Warn Icon"></td>
							<td>Needed (But not installed yet)</td>
"@
							}
							elseif($StatusID -eq 4){
								Add-Content $FilePath -Value @"
						<!--<td>&#8505 Pending (Not downloaded or installed yet)</td>-->
							<td><img class="icon" src="Media/Icons/warning.png" alt="Warn Icon"></td>
							<td>Pending (Not downloaded or installed yet)</td>
"@
							}
							elseif($StatusID -eq 5){
								Add-Content $FilePath -Value @"
						<!--<td>&#10060 Failed Installation</td>-->
							<td><img class="icon" src="Media/Icons/error.png" alt="Error Icon"></td>
							<td>Failed Installation</td>
"@
							}
							elseif($StatusID -eq 6){
								Add-Content $FilePath -Value @"
						<!--<td>&#9888 No Status Reported</td>-->
							<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"></td>
							<td>No Status Reported</td>
"@
							}
							else{
								Add-Content $FilePath -Value @"
						<!--<td>&#9888 No Status Reported</td>-->
						<td><img class="icon" src="Media/Icons/info.png" alt="Info Icon"></td>
						<td>No Status Reported</td>
"@
							}

							Add-Content $FilePath -Value @"
						<td>$UpdateName</td>
						<td class="Url"><a href="$SupportURL" target="_blank">$UpdSupportURL</a></td>
						</tr>
"@
						}
						
					}
					else {
						Add-Content $FilePath -Value @"
						<tr>
						<td>No Data available</td>
						</tr>
"@
					}
				}
				else {
					Add-Content $FilePath -Value @"
					<tr>
					<td>No Data available</td>
					</tr>
"@
				}
				
				#Finish Table
				Add-Content $FilePath -Value @"
    </tbody>
    </table>
"@
			}

		Clear-Host

		Write-Host -ForegroundColor Cyan "
    +----+ +----+     
    |####| |####|     
    |####| |####|       WW   WW II NN   NN DDDDD   OOOOO  WW   WW  SSSS
    +----+ +----+       WW   WW II NNN  NN DD  DD OO   OO WW   WW SS
    +----+ +----+       WW W WW II NN N NN DD  DD OO   OO WW W WW  SSS
    |####| |####|       WWWWWWW II NN  NNN DD  DD OO   OO WWWWWWW    SS
    |####| |####|       WW   WW II NN   NN DDDDD   OOOO0  WW   WW SSSS
    +----+ +----+       
"

		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "              Report Information for the last $days days"
		Write-Host "-----------------------------------------------------------------------------------"
		Write-Host "
    + Report Version           $UpdateStatusReportVersion
"
		Write-Host "`n-----------------------------------------------------------------------------------"
		Write-Host -ForegroundColor Yellow "              WSUS Updates Summary"
		
		# Display the summary at the end
		# Final Summary Output: Per Classification across all groups
		# Group summaries already collected in $classificationSummaries
		# So we just group by classification title and sum values
		
		$classificationSummaries |
			Group-Object ClassName |
			ForEach-Object {
				$classGroup = $_.Group

				$totals = @{
					TotalUpdates     = ($classGroup | Measure-Object -Property TotalUpdates -Sum).Sum
					Installed        = ($classGroup | Measure-Object -Property Installed -Sum).Sum
					NotApplicable    = ($classGroup | Measure-Object -Property NotApplicable -Sum).Sum
					Downloaded       = ($classGroup | Measure-Object -Property Downloaded -Sum).Sum
					NotInstalled     = ($classGroup | Measure-Object -Property NotInstalled -Sum).Sum
					Pending          = ($classGroup | Measure-Object -Property Pending -Sum).Sum
					Failed           = ($classGroup | Measure-Object -Property Failed -Sum).Sum
					NoStatus         = ($classGroup | Measure-Object -Property NoStatus -Sum).Sum
				}

				Write-Host -ForegroundColor Cyan "`n    + Classification    $($_.Name)"
				Write-Host "        Total Updates:                        $($totals.TotalUpdates)"
				Write-Host "        Installed:                            $($totals.Installed)"
				Write-Host "        Not Applicable:                       $($totals.NotApplicable)"
				Write-Host "        Downloaded:                           $($totals.Downloaded)"
				Write-Host "        Not Installed:                        $($totals.NotInstalled)"
				Write-Host "        Pending:                              $($totals.Pending)"
				Write-Host "        Failed:                               $($totals.Failed)"
				Write-Host "        No Status:                            $($totals.NoStatus)"
			}

		$classificationSummaries |
			Group-Object ClassName |
			ForEach-Object {
				$className = $_.Name  # Save outer loop variable
				Write-Host "`n-----------------------------------------------------------------------------------"
				Write-Host -ForegroundColor Cyan "              Classification $($className):"
				if($updSummaries){
					$filteredUpdates = $updSummaries | Where-Object { $_.ClassName -eq $className }
					
					if ($filteredUpdates.Count -gt 0) {
						$filteredUpdates | Select-Object UpdateStatus,UpdateName,SupportURL | Sort-Object UpdateName | Format-Table -AutoSize
					}
					else {
						Write-Host "None"
					}
				}
				else {
					Write-Host "None"
				}
			}
		
		
		Write-Progress -id 1 -Activity "Generating Updates Status Report" -Status "Finalizing:" -PercentComplete 100
	}
	
	function SelectHTMLReportFile {	
		$Options = @()
		$Input = ""
		$HTMLReportFileList = @()  
		$HTMLReportFileSelected = $false

		$GetHTMLReportFiles = Get-ChildItem -File "$FileDir" -Name -Include *.html
		$SumOfHTMLReportFile = $GetHTMLReportFiles.Count

		while (-not $HTMLReportFileSelected) {
			foreach ($HTMLReportFile in $GetHTMLReportFiles) {
				$HTMLReportFileList += $HTMLReportFile
			}

			Write-Host "`n Select a Report File:"
			for (($x = 0); $x -lt $SumOfHTMLReportFile; $x++) {
				Write-Host "    $x) $($HTMLReportFileList[$x])"
				$Options += $x
			}

			$Input = Read-Host "`n Please select an option (0/1/..)"
			Write-Log " User Input: $Input"

			if ($Options -contains [int]$Input) {
				$HTMLReportFileSelected = $true
				$GetFile = $HTMLReportFileList[$Input]
				$global:HTMLReportFileName = $GetFile
				Write-Host "`n Selected HTML Report File: $global:HTMLReportFileName" -ForegroundColor Yellow
			}
			else {
				Write-Host "`n Wrong Input" -ForegroundColor Red
				Write-Log " Wrong Input."
			}
		}
	}
	
	function ConvertToPDF {

		#$reportname = Get-ChildItem -Path C:\_psc\WSUS Files\WSUS_Reporting\Exports\*.html -Name
		#$reportname = $global:HTMLReportFileName
		$global:pdfPath = "$FileDir$htmlfilename.pdf"
		
		$msedge = @(
			"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
			"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
		)
		$msedge = $msedge | Where-Object { Test-Path $_ } | Select-Object -First 1
		
		if (-not $msedge) {
			Write-Log "ERROR: Microsoft Edge not found. Please install Microsoft Edge."
			Write-Host -ForegroundColor Red "Microsoft Edge not found. Please install Microsoft Edge."
			return
		}
		else{
		
			Write-Log "Converting HTML Report to PDF"
			Write-Log "Selected File: $htmlPath"
			Write-Log "Converted File: $global:pdfPath"

			Start-Sleep 5
		
			$arguments = "--print-to-pdf=""$global:pdfPath"" --headless --page-size=A4 --disable-extensions --no-pdf-header-footer --disable-popup-blocking --run-all-compositor-stages-before-draw --disable-checker-imaging --landscape ""file:///$htmlPath"""
			Write-Log "Edge arguments: $arguments"
			
			try{
			
				#Start-Process "msedge.exe" -ArgumentList @("--headless","--print-to-pdf=""$global:pdfPath""", "--landscape", "--page-size=A4","--disable-extensions","--no-pdf-header-footer","--disable-popup-blocking","--run-all-compositor-stages-before-draw","--disable-checker-imaging", "file:///$htmlPath") -Wait
				Start-Process -FilePath "$msedge" -ArgumentList $arguments -Wait -NoNewWindow
			
			}
			catch{
				Write-Log "ERROR: Convering to pdf failed. Reason: $_"
				Write-Host -ForegroundColor yellow "ERROR: Convering to pdf failed. Reason: $_"
			}
			Start-Sleep 5
		}
	}
	
	function GenerateHTMLReport {
    
		#Wait 10 Seconds - System needs to start background services etc. after foregoing reboot.
		Write-Progress -id 1 -Activity "Generating WSUS Updates Report" -Status "Generating Report:" -PercentComplete 0
		Start-Sleep -Seconds 10

		If (-Not ( Test-Path $RootDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'psc_wsusreporting'."
			New-Item -Path "C:\_psc\" -Name "psc_wsusreporting" -ItemType "directory"
		}

		If (-Not ( Test-Path $FileDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Exports'."
			New-Item -Path $RootDir -Name "Exports" -ItemType "directory"
		}

		If (-Not ( Test-Path $MediaDir ))
		{
			#Create Directory
			Write-Log "Create Directory 'Media'."
			New-Item -Path $FileDir -Name "Media" -ItemType "directory"
		}

		#Create HTML Report
		Write-Log "Create HTML Report."
		If (-Not ( Test-Path $FilePath ))
		{
			
			#Copy CSS Stylesheet and Images from DeploymentShare
			#Write-Log "Copy CSS Stylesheet and Images from DeploymentShare."
			#Copy-Item -Path "$global:FileShare\*" -Destination $global:MediaDir -Recurse -Force
			
			#Create File
			Write-Log "Create File."
			New-Item $FilePath -ItemType "file" | out-null
		
			#Add Content to the File
			Write-Log "Add Content to the File."
			Add-Content $FilePath -Value @"
<!doctype html>
<html>
	<head>
		<title>WSUS Update Report for $Hostname</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="description" content="this is a Report of updates that this windows update server is providing."/>
		<meta name="thumbnail" content=""/>
		<link rel="stylesheet" href="Media/styles.css" />
		<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
		<script>google.charts.load('current', {'packages':['corechart']});</script>
	</head>
	<body>
	<div id="main">
		<div id="title">
			<img id="default_logo" src="Media/Powershell_logo.png" alt="Logo">
			<h1 id="title">WSUS Update Report for $Hostname</h1>
            <table class="report-info">
			<tbody>
			<tr>
			<td>Report Template Version:</td>
			<td>Version $UpdateStatusReportVersion</td>
			</tr>
			</tbody>
			</table>
        </div>
"@
		}

		###### Function Calls ######
		Start-Update-Reporting

		#Finish HTML Report
		Add-Content $FilePath -Value @"
</div>
</body>
<footer>
</footer>
</html>
"@

	}
	
	<#
	# Starting Script
	#>
	Clear-Host
	GenerateHTMLReport
	SelectHTMLReportFile
	# Get HTML Report information
	$htmlfilename = $global:HTMLReportFileName -split ".html" | select-object -First 1
	$htmlPath = "$FileDir$global:HTMLReportFileName"
		
	
	do {
		$choice = Read-Host "`n Do you want to convert a report to pdf? (y/n)"
		Write-Log " User Input: $choice"
		switch ($choice) {
			"y" { ConvertToPDF }
			"n" {  }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne "y" -and $choice -ne "n")
	
	<#
	# Finalizing
	#>
	#Write-Progress -id 1 -Activity "Generating WSUS Synchronization Report" -Status "Finalizing:" -PercentComplete 98
	Write-Host "`n"
	
	if($global:pdfPath){
		Write-Host " PDF Report Location:    $global:pdfPath"
	}
	else{
		Write-Host " PDF Report Location:    No PDF File available"
	}
	if($htmlPath){
		Write-Host " HTML Report Location:   $htmlPath"
		Write-Host " 
    Automatically opening HTML Report now...
-----------------------------------------------------------------------------------"

		Invoke-Item "$htmlPath"
	}
	else{
		Write-Host " HTML Report Location:   No HTML File available"
	}
	
	
	Write-Host "
    Actions:
	
    e) Return to main menu"

    
	do {
		$choice = Read-Host " Choose an Option"
		Write-Log " User Input: $choice"
		switch ($choice) {
			#1 { GenerateHTMLReport }
			e { Show-Menu }
			default { 
				Write-Log " Wrong Input."
				Write-Host "Wrong Input. Please choose an option above." 
			}
		}

	} while ($choice -ne 1 -and $choice -ne "e")
	
	# Finish logging	
	Write-Log "Finish 'Update Status Report' Logging."
}

####
#### Main Menu Selection
####
Show-Menu

