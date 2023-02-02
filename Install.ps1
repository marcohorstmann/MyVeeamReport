#requires -Version 5.0
<#

    .DESCRIPTION
    Script to install/update MyVeeamReport

    .NOTES
    Author: Bernhard Roth
    Last Updated: 2 Februar 2023
    Version: 1.1


#>

# Customize to your requirements...
$Script = "C:\Temp\MyVeeamReport.ps1"
$Config = "C:\Temp\MyVeeamReport_config.ps1"
$Schedule = "C:\Temp\Schedule.ps1"


$URL_Script = "https://raw.githubusercontent.com/marcohorstmann/MyVeeamReport/main/MyVeeamReport.ps1"
$URL_Config = "https://raw.githubusercontent.com/marcohorstmann/MyVeeamReport/main/MyVeeamReport_config.ps1"
$URL_Schedule = "https://raw.githubusercontent.com/marcohorstmann/MyVeeamReport/main/Schedule.ps1"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# download latest version of script
Write-Host "Downloading latest version of script..."
try {
    Invoke-WebRequest -outfile $Script -uri $URL_Script
    Write-Host "The file [$Script] has been created."
} catch {
    throw $_.Exception.Message
}


# if the config file does not exist, create it.
if (-not(Test-Path -Path $Config -PathType Leaf)) {
    Write-Host "Downloading latest version of configuration file..."
    try {
        Invoke-WebRequest -outfile $Config -uri $URL_Config
        Write-Host "The file [$Config] has been created."
    } catch {
        throw $_.Exception.Message
    }
 }


# download schedule script
Write-Host "Downloading latest version of schedule script..."
try {
    Invoke-WebRequest -outfile $Schedule -uri $URL_Schedule
    Write-Host "Executing script..."
    Invoke-Expression -Command $Schedule
    Remove-Item $Schedule
    Write-Host "Scheduler set"
} catch {
    throw $_.Exception.Message
}
