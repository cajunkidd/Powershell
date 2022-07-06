<# 
.TITLE
    QC Automation Script

.FILENAME
    QCAutomation.ps1

.DESCRIPTION
    Baselines machines. Completing most of the steps on the QC checklist

.AUTHOR
    Tyler Neely 

.LASTMODIFIED
    7/5/2022

.NOTES
    -Rough Draft 
    -Software installs/uninstalls are using removeable D: drive for installers
    -Will clean up and add pauses for better flow control
    -Need checks for Mcafee | Office365 | NinjaRMM
#>
Write-Host -ForegroundColor Green "STARTING QC PROCESS"
Import-Module -Name PSWindowsUpdate -Force -InformationAction SilentlyContinue
$NewMachineName = Read-Host "Enter Computer Name. Current Computer Name: $env:COMPUTERNAME"
$InstalledApps  = @()
$MissingApps    = @()
Rename-Computer -NewName "$NewMachineName"

#SoftwareManagement
#$UnwantedApps = @()
Get-AppxPackage *skypeapp* | Remove-AppxPackage -ErrorAction SilentlyContinue
Get-AppxPackage *xbox* | Remove-AppxPackage -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
Get-AppxPackage *yourphone* | Remove-AppxPackage -ErrorAction SilentlyContinue
Get-AppxPackage *solitairecollection* | Remove-AppxPackage -ErrorAction SilentlyContinue
Get-AppxPackage *zune* | Remove-AppxPackage -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3
Write-Host -ForegroundColor Yellow "Installing Office365...PLEASE WAIT"
Start-Process -Wait -FilePath D:\QCAutomation\UninstallString.bat | Write-Error -Message "ERROR DURING INSTALL OR NOT INSTALLED"
Start-Process -Wait -FilePath D:\QCAutomation\AcroRdrDC2200120142_en_US.exe -ArgumentList '/sAll'
Start-Process -Wait -FilePath 'D:\QCAutomation\Firefox Setup 102.0.exe' -ArgumentList '/s' -PassThru
Start-Process -Wait -FilePath D:\ODT\O365std_x64.bat 

#WindowsUpdate
Write-Host -ForegroundColor Yellow "UPDATING WINDOWS...PLEASE WAIT"
Install-WindowsUpdate -ForceDownload -ForceInstall

#SoftwareChecks
Write-Host -ForegroundColor Yellow "Checking for Installed Software"
if(Test-Path 'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe'){
    Write-Host -ForegroundColor Green "ADOBE READER IS INSTALLED"
    (Get-ItemProperty -Path 'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe').VersionInfo.ProductVersion
    $InstalledApps += "Adobe"
}
else{
    Write-Warning "ADOBE READER IS NOT INSTALLED"
    $MissingApps += "Adobe"
}

if(test-path 'C:\Program Files\Mozilla Firefox\firefox.exe'){
    Write-Host -ForegroundColor Green "FIREFOX IS INSTALLED"
    (Get-ItemProperty -Path 'C:\Program Files\Mozilla Firefox\firefox.exe').VersionInfo.ProductVersion
    $InstalledApps += ",Firefox"
}
else{
    Write-Warning "FIREFOX IS NOT INSTALLED"
    $MissingApps += ",Firefox"
}

if(Test-Path 'C:\Program Files\Google\Chrome\Application\chrome.exe'){
    Write-Host -ForegroundColor Green "CHROME IS INSTALLED"
    (Get-ItemProperty -Path 'C:\Program Files\Google\Chrome\Application\chrome.exe').VersionInfo.ProductVersion
    $InstalledApps += ",Chrome"
}
else{
    Write-Warning "CHROME IS NOT INSTALLED"
    $MissingApps += ",Chrome"
}

#End
Write-Host -ForegroundColor Cyan "Checks Complete   Cleaning Up!"
cleanmgr.exe /verylowdisk
Remove-Item C:\users\nnadmin\Downloads\* -Force
Clear-RecycleBin -Force
Get-ComputerInfo -Property * | Format-List CsModel, BiosSeralNumber, CsName
Write-Host -ForegroundColor Yellow "Document Info Above on QC Checklist"
Write-Host -ForegroundColor Green "The following apps are installed: $InstalledApps"
Write-Host -ForegroundColor Red "The following apps are missing: $MissingApps"
Write-Warning "FININSHED! RESTART COMPUTER"
Restart-Computer -Confirm 

