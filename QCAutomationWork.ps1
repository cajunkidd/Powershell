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
$O365UninstallLocation = "D:\ODT\UninstallString.bat"
$O365InstallLocation   = "D:\ODT\O365std_x64.bat"
$NewMachineName        = Read-Host "Enter Computer Name:"
$UnnecessaryApps       = @('*skypeapp*' , '*xbox*' , '*yourphone*' , '*solitairecollection*' , '*zune*', '*mcafee*' )
$InstalledApps         = @()
$MissingApps           = @()
$SoftwareInfo          = [PSCustomObject]@{
    ComputerModel         = Get-ComputerInfo -Property CsModel
    ComputerSerialNumber  = Get-ComputerInfo -Property BiosSeralNumber
    NinjaInstalled        = $False
    NinajaVersion         = $null
    MalwarebytesInstalled = $False
    MalwarebytesVersion   = $null
    CoNNectInstalled      = $False
    CoNNectVersion        = $null
    AdobeInstalled        = $False
    AdobeVersion          = Get-ItemProperty -Path $AdobeLocation.ProductVersion
    ChromeInstalled       = $False
    ChromeVersion         = $null
    FirefoxInstalled      = $False
    FirefoxVersion        = $null
    UninstalledApps       = $UninstalledApps
    MissedBloatware       = $MissedBloatware
}

Import-Module -Name PSWindowsUpdate -Force -InformationAction SilentlyContinue
Rename-Computer -NewName "$NewMachineName"

#SoftwareUninstallandInstall
foreach($App in $UnnecessaryApps){
    Get-AppxPackage $App | Remove-AppxPackage -ErrorAction SilentlyContinue -InformationAction SilentlyContinue
    Get-AppxPackage $App | $UninstalledApps += $App
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
#WindowsUpdate
Write-Host -ForegroundColor Yellow "UPDATING WINDOWS...PLEASE WAIT"
Install-WindowsUpdate -ForceDownload -ForceInstall
#End
Write-Host -ForegroundColor Cyan "Checks Complete   Cleaning Up!"
cleanmgr.exe /verylowdisk
Remove-Item C:\users\nnadmin\Downloads\* -Force
Clear-RecycleBin -Force
Write-Warning "FININSHED! RESTART COMPUTER"
Restart-Computer -Confirm 

