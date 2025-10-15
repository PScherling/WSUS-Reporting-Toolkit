<#
.SYNOPSIS
	Automated (MDT/WDS) install of the PSC WSUS Reporting Tool from a deployment share.
.DESCRIPTION
    `custom_Install-PSC_wsus-report-tool.ps1` installs the PSC WSUS Reporting Tool on a target
    system by pulling its payload from a deployment share and performing post-copy setup.

    What it does:
      - Copies all files from: \\<MDT_FileSrv>\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data
        to: C:\_it\psc_wsusreporting  (preserves subfolders)
      - Creates the module path: C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting
      - Copies module files (.psm1 / .psd1) into the module path
      - Copies launcher: C:\Windows\System32\psc_wsusreporting.cmd
      - Imports the `psc_wsusreporting` module
      - Creates a desktop shortcut: “EF_WSUS-Report-Tool.lnk” → launch_psc_wsusreporting.bat
      - Writes a timestamped installer log locally and uploads it to the deployment log share

    Operator experience:
      - Uses `$PSScriptRoot`-independent network source with progress output
      - Robust try/catch with explicit step logging
      - Local log at: C:\_it\Configure_psc_wsusreporting_<Computer>_<DateTime>.log
        and upload to: \\<MDT_FileSrv>\Logs$\Custom\Configuration

	Requirements:
    - Run as Administrator.
    - Network connectivity and permissions to:
        - \\<MDT_FileSrv>\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data
        - \\<MDT_FileSrv>\Logs$\Custom\Configuration
    - Write access to:
        - C:\_it\
        - C:\Windows\System32\
        - C:\Program Files\WindowsPowerShell\Modules\
    - PowerShell 5.1+ (or PowerShell 7.x on Windows).	
	
.LINK
    https://learn.microsoft.com/windows/win32/wua_sdk/windows-update-agent--wua--api-reference
    https://learn.microsoft.com/windows-server/administration/windows-server-update-services/
	https://github.com/PScherling
	
.NOTES
          FileName: custom_Install-PSC_wsus-report-tool.ps1
          Solution: PSC_WSUS_Reporting (Automated)
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-05-28
          Modified: 2025-05-28

          Version - 0.0.1 - () - Finalized functional version 1.
          

          TODO:
		  	

.OUTPUTS
    Console output and log file:
        Local   : C:\_it\Configure_psc_wsusreporting_<COMPUTERNAME>_<YYYY-MM-DD_HH-mm-ss>.log
        Remote  : \\<MDT_FileSrv>\Logs$\Custom\Configuration\<same name>.log
		
.Example
	PS C:\> .\custom_Install-PSC_wsus-report-tool.ps1
    Pulls the WSUS Reporting Tool from the deployment share, installs the module, creates a desktop shortcut, imports the module, and uploads the installation log.
#>
function Start-Configuration {
    # Variables
    $user = "wdsuser"
    $pass = "Password"
	$FileSrv = "0.0.0.0" # MDT Server IP-Address
	
    $dest = "C:\_it\psc_wsusreporting"
    $source = "\\$($FileSrv)\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data"
	# Get all files and subdirectories from the source folder, including hidden/system files
	$items = Get-ChildItem -Path $source -Recurse
	
    $step1 = "false"
    $step2 = "false"
    $step3 = "false"
    $step4 = "false"
    $step5 = "false"
    $step6 = "false"
	
	# Log file path and function to log messages
	$config = "psc_wsusreporting"
	$CompName = $env:COMPUTERNAME
	$DateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
	$logFileName = "Configure_$($config)_$($CompName)_$($DateTime).log"

	$logFilePath = "\\$($FileSrv)\Logs$\Custom\Configuration"
	$logFile = "$($logFilePath)\$($logFileName)"

	$localLogFilePath = "C:\_it"
	$localLogFile = "$($localLogFilePath)\$($logFileName)"
	
	# Function to log messages with timestamps
    function Write-Log {
        param (
            [string]$Message
        )
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] $Message"
        #Write-Output $logMessage
        $logMessage | Out-File -FilePath $localLogFile -Append
    }
    
    # Start logging
    Write-Log "Start Logging."
    
    try {
        Write-Log "Step 1 - Download Files to target system"
        
        ### Step 1 - Download Files to target system
        if ($step1 -eq "false") {
            If (-Not (Test-Path $dest)) {
                # Create Directory
                try {
                    Write-Log "Creating target directory: $dest"
                    New-Item -Path "C:\_it\" -Name "psc_wsusreporting" -ItemType "directory"
                } catch {
                    Write-Log "Error: Directory can not be created: $_"
                    break
                }
            }

            try {
                Write-Log " Downloading files from $source to $dest"
                if (Test-Path $source) {
                    <#copy-item "$sourceFiles" -Destination "$dest\"
                    for ($i = 0; $i -le 100; $i = $i + 10) {
                        Write-Progress -Activity "File download in Progress" -Status "File Download Progress $i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 250
                    }#>
					
					# Copy each item individually
					$i = 0
					foreach ($item in $items) {
						# Ensure the target folder structure is created
						$targetItemPath = Join-Path -Path $dest -ChildPath $item.FullName.Substring($source.Length)
						
						# If it's a directory, create it
						if ($item.PSIsContainer) {
							if (-not (Test-Path -Path $targetItemPath)) {
								Write-Host "Creating directory: $targetItemPath"
								Write-Log " Creating directory: $targetItemPath"
								New-Item -Path $targetItemPath -ItemType Directory
							}
						}
						else {
							# If it's a file, copy it
							Write-Host "Download file: $item.FullName to $targetItemPath"
							Write-Log " Download file: $item.FullName to $targetItemPath"
							Copy-Item -Path $item.FullName -Destination $targetItemPath -Force
						}
						
						#Progress Bar
						if($i -lt 100){
							$i = $i + 10
							Write-Progress -Activity "File download in Progress" -Status "File Download Progress $i% Complete:" -PercentComplete $i
							Start-Sleep -Milliseconds 250
						}
						
					}
					
                    Write-Log "File download completed."
                } else {
                    Write-Log "Warning: Source path not found: $source"
                    break
                }
            } catch {
                Write-Log "Error: Files cannot be downloaded: $_"
                break
            }
            $step1 = "true"
        }

        Write-Log "Step 2 - Create Directories for our new module"
        
        ### Step 2 - Create Directories for our new module
        if ($step1 -eq "true") {
            if (-Not (Test-Path "C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting")) {
                try {
                    Write-Log "Creating module directory at 'C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting'"
					Write-Host " Creating module directory at 'C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting'"
                    New-Item -Path "C:\Program Files\WindowsPowerShell\Modules\" -Name "psc_wsusreporting" -ItemType "directory"
                } catch {
                    Write-Log "Error: Directory creation for module failed: $_"
					Write-Warning " Error: Directory creation for module failed: $_"
                    break
                }
            } else {
                Write-Log "Nothing to do. Directory already exists."
				Write-Host " Nothing to do. Directory already exists."
            }
            $step2 = "true"
        }

        Write-Log "Step 3 - Copy downloaded files to target directories"
        
        ### Step 3 - Copy downloaded files to target directories
        if ($step2 -eq "true") {
            if (-Not (Test-Path "C:\Windows\System32\psc_wsusreporting.cmd")) {
                try {
                    Write-Log "Copying psc_wsusreporting.cmd to System32"
					Write-Host " Copying psc_wsusreporting.cmd to System32"
                    copy-item "C:\_it\psc_wsusreporting\psc_wsusreporting.cmd" -Destination "C:\Windows\System32\"
                    for ($i = 0; $i -le 100; $i = $i + 10) {
                        Write-Progress -Activity "File copy in Progress" -Status "File Copy Progress $i% Complete:" -PercentComplete $i
                        Start-Sleep -Milliseconds 250
                    }
                    Write-Log "File copy completed for psc_wsusreporting.cmd."
					Write-Host " File copy completed for psc_wsusreporting.cmd."
                } catch {
                    Write-Log "Error: File copy failed for psc_wsusreporting.cmd: $_"
					Write-Warning " Error: File copy failed for psc_wsusreporting.cmd: $_"
                    break
                }
            } else {
                Write-Log "Nothing to do. psc_wsusreporting.cmd already exists."
				Write-Host " Nothing to do. psc_wsusreporting.cmd already exists."
            }

            # Repeat for other files (psc_wsusreporting.psm1, psc_wsusreporting.psd1)
            if (-Not (Test-Path "C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\psc_wsusreporting.psm1")) {
                try {
                    Write-Log "Copying psc_wsusreporting.psm1 to PowerShell Modules"
					Write-Host " Copying psc_wsusreporting.psm1 to PowerShell Modules"
                    copy-item "C:\_it\psc_wsusreporting\psc_wsusreporting.psm1" -Destination "C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\"
                } catch {
                    Write-Log "Error: File copy failed for psc_wsusreporting.psm1: $_"
					Write-Warning " Error: File copy failed for psc_wsusreporting.psm1: $_"
                    break
                }
            }

            if (-Not (Test-Path "C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\psc_wsusreporting.psd1")) {
                try {
                    Write-Log "Copying psc_wsusreporting.psd1 to PowerShell Modules"
					Write-Host " Copying psc_wsusreporting.psd1 to PowerShell Modules"
                    copy-item "C:\_it\psc_wsusreporting\psc_wsusreporting.psd1" -Destination "C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\"
                } catch {
                    Write-Log "Error: File copy failed for psc_wsusreporting.psd1: $_"
					Write-Warning " Error: File copy failed for psc_wsusreporting.psd1: $_"
                    break
                }
            }

            $step3 = "true"
        }

        Write-Log "Step 4 - Import the new module"
        
        ### Step 4 - Import our new module on the system
        if ($step3 -eq "true") {
            try {
                Write-Log "Importing psc_wsusreporting module"
				Write-Host " Importing psc_wsusreporting module"
                Import-Module psc_wsusreporting
            } catch {
                Write-Log "Error: Import of module failed: $_"
				Write-Warning " Error: Import of module failed: $_"
                break
            }
            $step4 = "true"
        }

        Write-Log "Step 5 - Desktop Icon psc_wsusreporting"
        
        ### Step 5 - Creating desktop shortcut
        if ($step4 -eq "true") {
			try{
				 Write-Log "Creating desktop shortcut for psc_wsusreporting"
				# Full path to your .bat file
				$batFilePath = "C:\_it\psc_wsusreporting\launch_psc_wsusreporting.bat"

				# Define the desktop shortcut path
				$desktop = [Environment]::GetFolderPath("Desktop")
				$shortcutPath = Join-Path $desktop "EF_WSUS-Report-Tool.lnk"

				# Create the shortcut
				$wsh = New-Object -ComObject WScript.Shell
				$shortcut = $wsh.CreateShortcut($shortcutPath)
				$shortcut.TargetPath = $batFilePath
				$shortcut.WorkingDirectory = Split-Path $batFilePath
				$shortcut.IconLocation = "C:\windows\System32\imageres.dll,298"
				$shortcut.Save()

				
			}
			catch{
				Write-Log "Error: Failed to create shortcut for psc_wsusreporting: $_"
				Write-Warning " Error: Failed to create shortcut for psc_wsusreporting: $_"
                break
			}
			Write-Log "Shortcut created on desktop: $shortcutPath"
            $step5 = "true"
        }

        if ($step5 -eq "true") {
            Write-Log "Configuration completed successfully."
			Write-Host "
-----------------------------------------------------------------------------------"
			Write-Host -ForegroundColor Green " Configuration completed successfully."
        }

    } catch {
        Write-Log "Error: An unexpected error occurred: $_"
		Write-Warning " Error: An unexpected error occurred: $_"
    } finally {
        # End logging
        Write-Log "Finish Logging."
    }
	
	<#
	# Finalizing
	#>
	
	# Upload logFile
	try{
		Copy-Item "$localLogFile" -Destination "$logFile"
	}
	catch{
		Write-Warning "ERROR: Logfile '$localLogFile' could not be uploaded to Deployment-Server.
		Reason: $_"
	}

	# Delete local logFile
	<#try{
		Remove-Item "$logFile" -Force
	}
	catch{
		Write-Warning "ERROR: Logfile '$logFile' could not be deleted.
		Reason: $_"
	}#>
}


Start-Configuration

