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

#Verify below paths to installers is correct!
$NinjaLocation         = "D:\QCAutomation\Ninja.exe"
$AdobeLocation         = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
$ChromeLocation        = "C:\Program Files\Google\Chrome\Application\chrome.exe"
$FirefoxLocation       = "C:\Program Files\Mozilla Firefox\firefox.exe"
$MalwareBytesLocation  = "C:\Program Files\Malwarebytes Endpoint Agent\UserAgent\Endpoint Agent Tray.exe"
$O365UninstallLocation = "D:\ODT\UninstallString.bat"
$O365InstallLocation   = "D:\ODT\O365std_x64.bat"
$NewMachineName        = Read-Host "Enter Computer Name:"
$UnnecessaryApps       = @('*skypeapp*' , '*xbox*' , '*yourphone*' , '*solitairecollection*' , '*zune*', '*mcafee*' )
$MissingApps           = @()
$SoftwareInfo = [PSCustomObject]@{
    ComputerModel         = Get-ComputerInfo -Property CsModel
    ComputerSerialNumber  = Get-ComputerInfo -Property BiosSeralNumber
    NinjaInstalled        = $False
    NinajaVersion         = $null
    MalwarebytesInstalled = $False
    MalwarebytesVersion   = (Get-ItemProperty -Path "C:\Program Files\Malwarebytes Endpoint Agent\UserAgent\Endpoint Agent Tray.exe").VersionInfo
    CoNNectInstalled      = $False
    AdobeInstalled        = $False
    AdobeVersion          = Get-ItemProperty -Path $AdobeLocation.ProductVersion
    ChromeInstalled       = $False
    ChromeVersion         = Get-ItemProperty $ChromeLocation.ProductVersion
    FirefoxInstalled      = $False
    FirefoxVersion        = Get-ItemProperty $FirefoxLocation.ProductVersion
    UninstalledApps       = $UninstalledApps
    MissedBloatware       = $MissedBloatware
}
Import-Module -Name PSWindowsUpdate -Force -InformationAction SilentlyContinue
Rename-Computer -NewName "$NewMachineName"
#SoftwareUninstallandInstall
foreach($App in $UnnecessaryApps){
    Get-AppxPackage $App | Remove-AppxPackage -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
    Get-AppxPackage $App += $MissedBloatware
    }
Write-Host -ForegroundColor Yellow "Installing Office & Updating Windows wait..."
Start-Process -Wait -FilePath D:\QCAutomation\UninstallString.bat | Write-Error -Message "ERROR - CHECK TO SEE IF ALREADY UNINSTALLED"
Install-WindowsUpdate -ForceDownload -ForceInstall
Start-Process -Wait -FilePath D:\ODT\O365std_x64.bat 
#Install ninja
if(Test-Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\NinjaRMMAgent){
    $SoftwareInfo.NinjaInstalled = $True
    (Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*NinjaRMMAgent*")
}
else{
    Write-Warning "Ninja Agent Not Installed"
    #instllninja
}
if(Test-Path -Path $MalwareBytesLocation){
    $SoftwareInfo.Malwarebytesinstalled = $True
}
if(test-path $AdobeLocation){
    $SoftwareInfo.AdobeInstalled = $true
    }
else{
    Start-Process -Wait -FilePath D:\QCAutomation\AcroRdrDC2200120142_en_US.exe -ArgumentList '/sAll'      
    }
if(Test-Path $FirefoxLocation){
    $SoftwareInfo.FirefoxInstalled = $True
    }
else{
    Start-Process -Wait -FilePath 'D:\QCAutomation\Firefox Setup 102.0.exe' -ArgumentList '/s' -PassThru
    }
if(Test-Path $ChromeLocation){
    $SoftwareInfo.ChromeInstalled = $True
    $SoftwareInfo.ChromeVersion = Get-ItemProperty -Path $ChromeLocation.Product
    }
#End
cleanmgr.exe /verylowdisk
Remove-Item C:\users\nnadmin\Downloads\* -Force
Clear-RecycleBin -Force
Write-Warning "FININSHED! RESTART COMPUTER"
Restart-Computer -Confirm 

