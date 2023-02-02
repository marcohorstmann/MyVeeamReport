#requires -Version 5.0
<#

    .DESCRIPTION
    Script to create a scheduled task to run MyVeeamReport

    .NOTES
    Author: Bernhard Roth
    Last Updated: 15 November 2022
    Version: 1.0

#>

# Customize to your requirements...
$Script = "C:\Temp\MyVeeamReport.ps1"
$WorkDir = "C:\Temp"

$User = "NT AUTHORITY\SYSTEM"

# Task trigger: Weekly
$Trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Wednesday -At 11pm
# Daily
# $Trigger = New-ScheduledTaskTrigger -Daily -At 11:30am

# Task action: Run Script as argument to PowerShell.exe in specified working directory
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $Script -WorkingDirectory $WorkDir

# Task settings: Win8/Server2016 compatibility, 15 minute max. runtime, start if schedule is missed (e.g. for updates or reboot)
$Setting = New-ScheduledTaskSettingsSet -Compatibility Win8 -ExecutionTimeLimit (New-TimeSpan -Minutes 15) -StartWhenAvailable
# Optional: Run only if network is available for email. Useful if report is not saved locally
# -RunOnlyIfNetworkAvailable

# Register Task
Register-ScheduledTask -TaskName "Execute Veeam Backup Report" -Trigger $Trigger -User $User -Action $Action -Settings $Setting -RunLevel Highest -Force