<#
.Synopsis
Used to install Micro Focus Content Manager 9.4
.Description
Used to install Micro Focus Content Manager 9.4
.Notes
Can be used to upgrade old versions or for fresh install.
Change parameters to match your enviroment before running.
Author: Andrew Kapaldo
#>

Write-Host "Installing Content Manager 9.4"

#Uninstall Old Version
Write-Host "Checking for and removing old versions..."
try {
Get-Package -Name "Content Manager Core Components x86" -MaximumVersion 9.31.300 -ErrorAction SilentlyContinue | Uninstall-Package
}
catch{
Write-Host "Max Version 9.31.300 does not exist."
}
Write-Host "Uninstalled old version."

#Copy MSI File to Local Machine 
#!!!!!!!! Update with file path  !!!!!!!!!
$path = $Null
$logfile = "" #Log path here
Write-Host "Copying Content Manager 9.4"
Copy-Item -Path $path -destination "c:\temp"
Write-Host "Copied Content Manager 9.4, Pausing for 5 seconds."
Start-Sleep -seconds 5

#Install CM from local file
Write-Host "Installing Content Manager 9.4"
$args = "/i `"C:\Temp\CM_Client_x86.msi`" /qn /l*vx $logfile"
Start-Process "msiexec.exe" -ArgumentList $args -Wait

#Clean up files
Write-Host "Cleaning up files."
if (Test-Path -Path "C:\Users\Public\Desktop\Content Manager Desktop.lnk"){
	Remove-Item -Path "C:\Users\Public\Desktop\Content Manager Desktop.lnk" -Force
	}
if (Test-Path -Path "C:\Temp\CM_Client_x86.msi"){
	Remove-Item -Path "C:\Temp\CM_Client_x86.msi" -Force
	}
if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Content Manager\Content Manager Desktop.lnk"){
	Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Content Manager\Content Manager Desktop.lnk" -Force
	}

#Add Web Links to Apps Folder
#!!!!!!!!!!!!! Add Links to Target Parameters !!!!!!!!!!!!!!!
$Icon = "C:\Program Files (x86)\Micro Focus\Content Manager\TRIM.exe"
$WshShell = New-Object -comObject WScript.Shell
$Target1 = $Null
$Target2 = $Null

if (Test-Path -Path "C:\Users\Public\Desktop\Apps\Web Client Prod.lnk") {
}
else {
$Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Apps\Web Client Prod.lnk")
$Shortcut.TargetPath = $Target1
$Shortcut.IconLocation = $Icon
$Shortcut.Save()
}

if (Test-Path -Path "C:\Users\Public\Desktop\Apps\Web Client Train.lnk") {
}
else {
$Shortcut1 = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Apps\Web Client Train.lnk")
$Shortcut1.TargetPath = $Target2
$Shortcut1.IconLocation = $Icon
$Shortcut1.Save()
}

#Test for completed install
if (Get-Package -Name "Content Manager Client Components x86" -MinimumVersion 9.42.602 -ErrorAction SilentlyContinue) {
	Write-Host "Content Manager 9.4 Installed."}
else {
	Write-Warning "Failed to install Content Manager 9.4" }
