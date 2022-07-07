$Count = 2
$Ping = @()
$csv_log = "C:\Users\vincent\Documents\test.csv"
$Computer = "www.google.com";
While ($Count -gt 0)
{
    $Ping += Get-WmiObject Win32_PingStatus -Filter "Address = '$Computer'" | 
    Select @{Label="TimeStamp";
    Expression={Get-Date}},@{Label="Source";Expression={ $_.__Server }},@{Label="Destination";
    Expression={ $_.Address }},IPv4Address,@{Label="Status";
    Expression={ If ($_.StatusCode -ne 0) {"Failed"
                                            [Console]::Beep()
                                           } 
                                           Else {""}
             }
 },ResponseTime
    $Count --

    $Ping | Select TimeStamp,Source,Destination,IPv4Address,Status,ResponseTime | Export-Csv -Append $csv_log -NoTypeInformation
    Clear-Variable -name Ping
    Start-Sleep -Seconds 3
}