# README - PowerShell WSUS Reporting
Statusberichte können über das PowerShell Modul "psc_wsusreporting" abgerufen und exportiert werden. Dieses kann manuell oder auch via WDS automatisch installiert.

In case you have or want to install the 'psc_wsusreporting' on your system after your manual installation of a Windows Server Edition you can do that with the following package.

To get the 'psc_wsusreporting' on your system, just do these simple steps:

1. Upload this zip file somewhere to your system, where you want to install it like to "D:\TEMP" and extract it.
2. Extract the zip file
3. Review the extracted contents. There should be a directory called 'Data' and a PowerShell Script for installation.
4. Logon to your server and enter the cli
5. In the new cli window:
  a) Review your systems execution policy by running the command 'Get-ExecutionPolicy'
    Get-ExecutionPolicy
    i) If the Policy is set to restricted, run this command 'Set-ExecutionPolicy Bypass'
      Set-ExecutionPolicy Bypass
    ii) If you don't change the policy to bypass, you are not allowed to execute the install script.
  b) Verify that the ExecutionPolicy is set to 'Bypass'
    Get-ExecutionPolicy
  c) Navigate into the directory where you have extracted the files (e.g. 'D:\Temp')
    cd d:\temp
 
  d) Run the install script by executing this command: '.\manual_Install-psc_wsus-report-tool.ps1'
    .\manual_Install-psc_wsus-report-tool.ps1
 
  Let the installer do it's thing

  e) After the installation has finished, you should see a message 'Configuration completed successfully.'
  f) Press any key to exit the script

  g) Run the command 'Set-ExecutionPolicy Default'
    Set-ExecutionPolicy Default
    Get-ExecutionPolicy
    Review that the policy is not set to bypass anymore

6. Now you can run the psc_wsusreporting slightly smiling face 
Via PowerShell: 
start psc_wsusreporting

or via Desktop Link.
