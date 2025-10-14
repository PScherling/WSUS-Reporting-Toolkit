<#
.SYNOPSIS
    Launches the PSC WSUS Reporting Tool.
	
.DESCRIPTION
    The `psc_wsusreporting.psm1` module exposes a single entry-point function, `psc_wsusreporting`,
    which starts the WSUS Reporting Tool UI by launching:
        C:\_it\psc_wsusreporting\psc_wsusreporting.ps1
    in a new, maximized PowerShell process with `-ExecutionPolicy Bypass`.

    Behavior:
      - Spawns a separate PowerShell session using $PSHOME\powershell.exe
      - Maximizes the window for better operator visibility
      - Surfaces launch errors to the console

    Use this module after the tool has been installed (manual or automated installer),
    so operators can simply run `psc_wsusreporting` from any PowerShell prompt.
	
.LINK
	https://learn.microsoft.com/windows-server/administration/windows-server-update-services/
    https://learn.microsoft.com/windows/win32/wua_sdk/windows-update-agent--wua--api-reference
    https://github.com/PScherling
    
.NOTES
          FileName: psc_wsusreporting.psm1
          Solution: PSC_WSUS_Reporting
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-05-02
          Modified: 2025-05-28

          Version - 0.0.1 - () - Finalized functional version 1.
          

          TODO:
		  
.REQUIREMENTS
    - PowerShell 5.1+ (or PowerShell 7.x on Windows)
    - File present at: C:\_it\psc_wsusreporting\psc_wsusreporting.ps1		  
		
.Example
	PS C:\> Import-Module psc_wsusreporting
    Imports the PSC WSUS Reporting Tool launcher module.

	PS C:\> psc_wsusreporting
    Opens the WSUS Reporting Tool in a new, maximized PowerShell window.
#>

function psc_wsusreporting {
	$pscwsusreporting = "C:\_it\psc_wsusreporting\psc_wsusreporting.ps1"
	if($pscwsusreporting){
		try{
			Start-Process -FilePath "$PSHOME\powershell.exe" -ArgumentList "-windowstyle maximized -ExecutionPolicy Bypass -File $pscwsusreporting" #-PassThru
		}
		catch{
			Write-Error " Can't launch PowerShell script:
			$_"
		}
	}
	else{
		Write-Error "PowerShell Script file not found:
		$_"
	}

}
