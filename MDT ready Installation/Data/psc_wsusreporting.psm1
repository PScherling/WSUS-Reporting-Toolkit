<#
.DESCRIPTION
    
.LINK
    
.NOTES
          FileName: psc_wsusreporting.psm1
          Solution: 
          Author: Patrick Scherling
          Contact: @Patrick Scherling
          Primary: @Patrick Scherling
          Created: 2025-05-02
          Modified: 2025-05-28

          Version - 0.0.1 - () - Finalized functional version 1.
          

          TODO:
		  
		
.Example
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