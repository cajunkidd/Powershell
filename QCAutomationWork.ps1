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
    Will clean up and add pauses for better flow control
    Verify below paths to installers is correct!
#>
Import-Module -Name PSWindowsUpdate -Force
$NinjaLocation         = "D:\QCAutomation\Ninja.exe"
$AdobeLocation         = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
$ChromeLocation        = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$FirefoxLocation       = "C:\Program Files\Mozilla Firefox\firefox.exe"
$NewMachineName        = Read-Host "Enter Computer Name"
$UnnecessaryApps       = @('*skypeapp*' , '*xbox*' , '*yourphone*' , '*solitairecollection*' , '*zune*', '*mcafee*' , '*disney*' )
$MissingApps           = @()
$OfficeBuild           = Get-ItemPropertyValue -Path HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration -Name “VersionToReport”
$SoftwareInfo = [PSCustomObject]@{
    ComputerModel         = Get-ComputerInfo -Property CsModel
    ComputerSerialNumber  = Get-ComputerInfo -Property BiosSeralNumber
    OfficeInstalled       = "False"
    OfficeBuildNum        = $null
    NinjaInstalled        = "False"
    MalwarebytesInstalled = "False"
    CoNNectInstalled      = "False"
    TeamviewerInstalled   = "False"
    ChromeInstalled       = "False"
    AdobeInstalled        = "False"
    FirefoxInstalled      = "False"
    NinajaVersion         = $null
    AdobeVersion          = (Get-ItemProperty -Path "$AdobeLocation").VersionInfo.ProductVersion
    MalwarebytesVersion   = (Get-ItemProperty -Path "C:\Program Files\Malwarebytes Endpoint Agent\UserAgent\Endpoint Agent Tray.exe").VersionInfo.ProductInfo
    ChromeVersion         = (Get-ItemProperty $ChromeLocation).VersionInfo.ProductVersion
    FirefoxVersion        = (Get-ItemProperty "$FirefoxLocation").VersionInfo.ProductVersion 
    UninstalledApps       = $MissingApps
    MissedBloatware       = $MissedBloatware
}
#SoftwareInstalls
#Import-Module -Name PSWindowsUpdate -Force
Rename-Computer -NewName "$NewMachineName"
foreach($App in $UnnecessaryApps){
   
    Get-AppxPackage $App | Remove-AppxPackage -ErrorAction SilentlyContinue -InformationAction SilentlyContinue

    }
Write-Host -ForegroundColor Yellow "INSTALLING OFFICE & UPDATING WINDOWS...WAIT"
Start-Process -Wait -FilePath D:\QCAutomation\UninstallString.bat | Write-Error -Message "ERROR - CHECK TO SEE IF ALREADY UNINSTALLED"
Start-Sleep -Seconds 69
Install-WindowsUpdate -ForceDownload -ForceInstall
Start-Process -Wait -FilePath D:\ODT\O365std_x64.bat 
Start-Sleep -Seconds 69
Start-Process -Wait -FilePath "D:\QCAutomation\businesshealthpartnersleesville-5.3.4287-windows-installer.msi" -ArgumentList "/qn" -PassThru

if($OfficeBuild -eq "16.0.15330.20230"){
    $SoftwareInfo.OfficeBuildNum = "Correct Version Installed: O365BusinessRetail"
    }
if(Test-Path -Path HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration){
    $SoftwareInfo.OfficeInstalled = "True"
    }
else{
    $SoftwareInfo.UninstalledApps += "Office"
    }
if(Test-Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*NinjaRMMAgent*){
    $SoftwareInfo.NinjaInstalled = "True"
    $SoftwareInfo.NinajaVersion = Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*NinjaRMMAgent*" | Select-Object -Property DisplayVersion
    }
else{
    Write-Warning "Ninja Agent Not Installed"
    $MissingApps += "Ninja"
    }
if(Test-Path -Path "C:\Program Files\Malwarebytes Endpoint Agent\UserAgent\Endpoint Agent Tray.exe"){
    $SoftwareInfo.Malwarebytesinstalled = "True"
    }
else{
    $MissingApps += "Malwarebytes"
    }
if(Test-Path "C:\Program Files (x86)\DeskDirector Client Portal\DeskDirectorPortal.exe"){
    $SoftwareInfo.CoNNectInstalled = "True"
    }
else{
    $MissingApps += "CoNNect"
    }
if(Test-Path "C:\Program Files (x86)\TeamViewer\TeamViewer.exe"){
    $SoftwareInfo.TeamviewerInstalled = "True"
    }
else{
    $MissingApps += "Teamviewer"
    }
if(test-path $AdobeLocation){
    $SoftwareInfo.AdobeInstalled = "True"
    }
else{
    Start-Process -Wait -FilePath D:\QCAutomation\AcroRdrDC2200120142_en_US.exe -ArgumentList '/sAll'      
    }
if(Test-Path $FirefoxLocation){
    $SoftwareInfo.FirefoxInstalled = "True"
    }
else{
    Start-Process -Wait -FilePath 'D:\QCAutomation\Firefox Setup 102.0.exe' -ArgumentList '/s' -PassThru
    }
if(Test-Path $ChromeLocation){
    $SoftwareInfo.ChromeInstalled = "True"
    }
else{
    $MissingApps += "Chrome"
    }

cleanmgr.exe /verylowdisk /autoclean
Remove-Item C:\users\nnadmin\Downloads\* -Force
Clear-RecycleBin -Force
Write-Warning "FININSHED! RESTART COMPUTER"
$SoftwareInfo | Format-List
Start-Sleep -Seconds 30
Restart-Computer -Confirm 