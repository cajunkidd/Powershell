<####################################>
<##### DEFRAG AND OPTIMIZE DRIVE ####>
<####################################>

Optimize-Volume -DriveLetter C -Defrag -Verbose 



Clear-RecycleBin -Force


<####################################>
<###### DUMP LOG INTO TEXT FILE #####>
<####################################>

<# GOING TO CIRCLE BACK ON THIS #>

<####################################>
<#### DOWNLOAD CCLEANER FROM WEB ####>
<####################################>

$source = 'https://www.ccleaner.com/ccleaner/download/ccsetup563.exe'
$destination = 'C:\Users\vincent\Downloads\ccsetup563.exe'
Invoke-WebRequest -Uri $source -Outfile $destination


<####################################>
<########### RUN CCLEANER ###########>
<####################################>

$cCleanerInstalled = Test-Path -Path 'C:\Users\vincent\Downloads\ccsetup563.exe'

If ($cCleanerInstalled){
    Write-Host "Installed - running the cleaner!"
    Start-Process -FilePath "C:\Program Files\CCleaner\CCleaner64.exe" /Auto /Shutdown
} ELSE {
    Write-Host "Running Installation Now!"
}

<####################################>
<####### DELETE TEMP FILES ##########>
<####################################>

Remove-Item $env:TEMP\* -Recurse -Force -ErrorAction SilentlyContinue

