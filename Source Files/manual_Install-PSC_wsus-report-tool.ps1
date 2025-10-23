<#
.SYNOPSIS
	Manually installs the PSC WSUS Reporting Tool and its PowerShell module on the local system.
.DESCRIPTION
    `manual_Install-PSC_wsus-report-tool.ps1` performs a local, manual installation of the
    PSC WSUS Reporting Tool without needing a deployment server.

    What it does:
      - Copies all tool files from the script’s local `.\Data` folder into `C:\_it\psc_wsusreporting`
      - Creates the module folder: C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting
      - Copies `psc_wsusreporting.psm1/.psd1` into the module folder
      - Copies `psc_wsusreporting.cmd` into `C:\Windows\System32` for easy shell launch
      - Imports the `psc_wsusreporting` module
      - Creates a desktop shortcut (WSUS-Report-Tool.lnk) targeting `launch_psc_wsusreporting.bat`
      - Writes timestamped logs to `C:\_it\<Configure_psc_wsusreporting_*.log>`

    The script is verbose and operator-friendly:
      - Uses `$PSScriptRoot` to resolve its own `.\Data` source folder
      - Provides progress messages and clear success/error output
      - Includes robust try/catch error handling and logging for each step
	  
.LINK
    https://learn.microsoft.com/windows/win32/wua_sdk/windows-update-agent--wua--api-reference
    https://learn.microsoft.com/windows-server/administration/windows-server-update-services/get-started/windows-server-update-services-wsus
	https://github.com/PScherling
	
.NOTES
          FileName: manual_Install-PSC_wsus-report-tool.ps1
          Solution: PSC_WSUS_Reporting
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-05-02
          Modified: 2025-05-09

          Version - 0.0.1 - () - Finalized functional version 1.
          

          TODO:

.REQUIREMENTS
    - Run as Administrator.
    - PowerShell 5.1+ (or PowerShell 7.x on Windows).
    - Sufficient permissions to write to:
        - C:\_it\
        - C:\Windows\System32\
        - C:\Program Files\WindowsPowerShell\Modules\

.OUTPUTS
    Console output and installer log:
        C:\_it\Configure_psc_wsusreporting_<COMPUTERNAME>_<YYYY-MM-DD_HH-mm-ss>.log
		
.Example
	PS C:\> .\manual_Install-PSC_wsus-report-tool.ps1
    Runs the local/manual installer using the script’s .\Data source folder.

	PS C:\> powershell.exe -ExecutionPolicy Bypass -File "C:\Install\manual_Install-PSC_wsus-report-tool.ps1"
    Executes the installer from a local path, bypassing execution policy for the session.
#>
function Start-Configuration {
    ### 
	### Variables
	###
	
	# Get the directory where the script is located
	$scriptDirectory = $PSScriptRoot
	
    $dest = "C:\_it\psc_wsusreporting"
    $source = "$scriptDirectory\Data"
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
	Write-Host "-----------------------------------------------------------------------------------
    Starting PSC WSUS Reporting Tool Installer...
-----------------------------------------------------------------------------------"
    
    try {
        Write-Log "Step 1 - Copy Files to target directory"
        Write-Host " Step 1 - Copy Files to target directory"
        ### Step 1 - Download Files to target directory
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
                Write-Log " Copy files from $source to $dest"
				Write-Host " Copy files from $source to $dest"
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
						#$i = $i + 10
						#Write-Progress -Activity "File download in Progress" -Status "File Download Progress $i% Complete:" -PercentComplete $i
						#Start-Sleep -Milliseconds 250
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
		Write-Host "
-----------------------------------------------------------------------------------"
        Write-Host " Step 2 - Create Directories for our new module"
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

        Write-Log "Step 3 - Copy files to target directories"
		Write-Host "
-----------------------------------------------------------------------------------"
        Write-Host " Step 3 - Copy files to target directories"
        ### Step 3 - Copy files to target directories
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
		Write-Host "
-----------------------------------------------------------------------------------"
        Write-Host " Step 4 - Import the new module"
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
		Write-Host "
-----------------------------------------------------------------------------------"

        ### Step 5 - Creating desktop shortcut
        if ($step4 -eq "true") {
			try{
				Write-Log "Creating desktop shortcut for psc_wsusreporting"
				# Full path to your .bat file
				$batFilePath = "C:\_it\psc_wsusreporting\launch_psc_wsusreporting.bat"

				# Define the desktop shortcut path
				$desktop = [Environment]::GetFolderPath("Desktop")
				$shortcutPath = Join-Path $desktop "WSUS-Report-Tool.lnk"

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
	Read-Host -prompt " Press any key to finish..."
}

Start-Configuration
