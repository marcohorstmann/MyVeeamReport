#requires -Version 5.0
<#

    .SYNOPSIS
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication.

    .DESCRIPTION
    My Veeam Report is a flexible reporting script for Veeam Backup and
    Replication. This report can be customized to report on Backup, Replication,
    Backup Copy, Tape Backup, SureBackup and Agent Backup jobs as well as
    infrastructure details like repositories, proxies and license status. Work
    through the User Variables to determine what you would like to see and
    determine if you would like to save the results to a file or have them
    emailed to you.

    .EXAMPLE
    .\MyVeeamReport.ps1
    Run script from (an elevated) PowerShell console

    .NOTES
    New Authors: Marco Horstmann, Bernhard Roth & Herbert Szumovski
    Last Updated: 20 Feburary 2023
    Version: 12.0.0.0
    New Authors: Marco Horstmann & Herbert Szumovski
    Last Updated: 23 March 2022
    Version: 11.0.1.4
    Original Author: Shawn Masterson
    Last Updated: December 2017
    Version: 9.5.3

    Requires:
    Veeam Backup & Replication v11.0 or later (full or console install)
    VMware Infrastructure

#>

#region User-Variables
. .\MyVeeamReport_config.ps1
#endregion

#region VersionInfo
$MVRversion = "12.0.0.1"

# Version 12.0.0.1 MH - 2023-02-21
# Changed tape code to show GFS media pools
# Fixed error with backup copy session reporting (Thx to Nathan)

# Version 12.0.0.0 MH - 2023-02-20
# Added some code to support file backup jobs
# Added support for immediate copy (tested with v12 Unified Backup Copy Job)

# Version 11.0.1.6 MH - 2023-02-04
# Fixed Bug with license code and NFR licenses

# Version 11.0.1.5 BR - 2022-11-15
# Update license retrieval code, support different license variants
# Put configuration in external file
# Add informational section header
# Add option to show unprotected VMs for informational purposes only
# Use computername instead of "localhost"
# Fix VBR version in report header
# Suppress error on Get-VBRTapeVault if not licensed
# Fix HTML validation errors
# Use PSScriptChecker to implement PS syntax and verb recommendations

# Version 11.0.1.4 MH - 2022-03-23
# Merged Herberts and my version

# Version 11.0.1.3 HS - 2022-03-20
# fixed backup copy reports

# Version 11.0.1.2 HS - 2022-03-14
# fixed Agent job status
# added non-veeam data column to repo display
# modified html output a little bit to enhance readability

# Version 11.0.1.1 HS - 2022-03-12
# Removed support for versions below 11, unfortunately I have no time to
# support the old Cmdlets
# fixed license details display
# fixed mediapool display
# fixed Repository display and added a column for backupsize
# added VCSP and SAN Snapshots as additional Repositorytypes
# added a few runtime messages
# added more detailed version display
# fixed next runtime display in job report


# Version 11.0.0.1 HS - 2021-01-27
# Support of V11 GA, removed powershell snapin, replaced deprecated calls

# Version 11.0.0 - MH
# Updated script to work with VBR 11 Beta

# Version 9.5.3 - SM
# Updated property changes introduced in VBR 9.5 Update 3

# Version 9.5.1.1 - SM
# Minor bug fixes:
# Removed requires VBR snapin
# Fixed HourstoCheck variable in Get-VMsBackupStatus function
# Version 9.5.1 - SM
# Updated HTML formatting - thanks for the inspiration Nick!
# Report header and email subject now reflect results (Failed/Warning/Success)
# Added report section - VMs Backed Up by Multiple Jobs within RPO
# Added report section - Repository Permissions for Agent Jobs
# Added Description field for Agent Job Status to identify type of Agent
# Added Next Run field for Agent Job Status (Fixed in VBR 9.5 Update 1)
# Added Next Run field for Configuration Backup Status (Fixed in VBR 9.5 Update 1)
# Added more details to VMs with No Successful/Successful/with Warnings within RPO
# Appended date and time to email attachment file name
# Added ability to append date and time to email subject
# Added ability to send email via SSL/TLS
# Renamed Endpoints to Agents
#
# Version 9.0.3 - SM
# Added report section - VM Backup Protection Summary (across entire infrastructure)
# Split report section - Split out VMs with only Backups with Warnings within RPO to separate from Successful
# Added report section - Backup Job Size (total)
# Added report section - All Backup Sessions
# Added report section - All Backup Tasks
# Added report section - Running Backup Tasks
# Added report section - Backup Tasks with Warnings or Failures
# Added report section - Successful Backup Tasks
# Added report section - Replication Job/Session Summary
# Added report section - Replication Job Status
# Added report section - All Replication Sessions
# Added report section - All Replication Tasks
# Added report section - Running Replication Jobs
# Added report section - Running Replication Tasks
# Added report section - Replication Job/Sessions with Warnings or Failures
# Added report section - Replication Tasks with Warnings or Failures
# Added report section - Successful Replication Jobs/Sessions
# Added report section - Successful Replication Tasks
# Added report section - Backup Copy Session Summary
# Added report section - Backup Copy Job Status
# Added report section - Backup Copy Job Size (total)
# Added report section - All Backup Copy Sessions
# Added report section - All Backup Copy Tasks
# Added report section - Idle Backup Copy Sessions
# Added report section - Pending Backup Copy Tasks
# Added report section - Working Backup Copy Jobs
# Added report section - Working Backup Copy Tasks
# Added report section - Backup Copy Sessions with Warnings or Failures
# Added report section - Backup Copy Tasks with Warnings or Failures
# Added report section - Successful Backup Copy Sessions
# Added report section - Successful Backup Copy Tasks
# Added report section - Tape Backup Session Summary
# Added report section - Tape Job Status
# Added report section - All Tape Backup Sessions
# Added report section - All Tape Backup Tasks
# Added report section - Waiting Tape Backup Sessions
# Added report section - Idle Tape Backup Sessions
# Added report section - Pending Tape Backup Tasks
# Added report section - Working Tape Backup Jobs
# Added report section - Working Tape Backup Tasks
# Added report section - Tape Backup Sessions with Warnings or Failures
# Added report section - Tape Backup Tasks with Warnings or Failures
# Added report section - Successful Tape Backup Sessions
# Added report section - Successful Tape Backup Tasks
# Added report section - All Tapes
# Added report section - All Tapes by (Custom) Media Pool
# Added report section - All Tapes by Vault
# Added report section - All Expired Tapes
# Added report section - Expired Tapes by (Custom) Media Pool - Thanks to Patrick IRVING & Olivier Dubroca!
# Added report section - Expired Tapes by Vault
# Added report section - All Tapes written to within time frame ($reportMode)
# Added report section - Endpoint Backup Job Size (total)
# Added report section - All Endpoint Backup Sessions
# Added report section - SureBackup Session Summary
# Added report section - SureBackup Job Status
# Added report section - All SureBackup Sessions
# Added report section - All SureBackup Tasks
# Added report section - Running SureBackup Jobs
# Added report section - Running SureBackup Tasks
# Added report section - SureBackup Sessions with Warnings or Failures
# Added report section - SureBackup Tasks with Warnings or Failures
# Added report section - Successful SureBackup Sessions
# Added report section - Successful SureBackup Tasks
# Added report section - Configuration Backup Status
# Added report section - Scale Out Repository Info - Thanks to Patrick IRVING & Olivier Dubroca!
# Added exclusion for Templates to VM Backup Protection sections
# Added Last Start and End times to VMs with Successful/Warning Backups
# Added Dedupe and Compression to Backup/Backup Copy/Replication session detailed info
# Added ability to report only on particular jobs (backup/replica/backup copy/tape/surebackup/endpoint)
# Added Mode/Type and Maximum Tasks to Proxy and Repository Info
# Filtered some heavy lifting commands to only run when/if needed
# Converted durations from Mins to HH:MM:SS
# Added html formatting of cells (vertical-align: middle;text-align:center;)
# Lots of misc tweaks/cleanup
#
# Version 9.0.2 - SM
# Fixed issue with Proxy details reported when using IP address instead of server names
# Fixed an issue where services were reported multiple times per server
#
# Version 9.0.1 - SM
# Initial version for VBR v9
# Updated version to follow VBR version (VeeamMajorVersion.VeeamMinorVersion.MVRVersion)
# Fixed Proxy Information (change in property names in v9)
# Rewrote Repository Info to use newly available properties (yay!)
# Updated Get-VMsBackupStatus to remove obsolete commandlet warning (Thanks tsightler!)
# Added ability to run from console only install
# Added ability to include VBR server in report title and email subject
# Rewrote License Info gathering to allow remote info gathering
# Misc minor tweaks/cleanup
#
# Version 2.0 - SM
# Misc minor tweaks/cleanup
# Proxy host IP info now always returns IPv4 address
# Added ability to query Veeam database for Repository size info
#   Big thanks to tsightler - http://forums.veeam.com/powershell-f26/get-vbrbackuprepository-why-no-size-info-t27296.html
# Added report section - Backup Job Status
# Added option to show detailed Backup Job/Session information (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB))
# Added report section - Running VM Restore Sessions
# Added report section - Completed VM Restore Sessions
# Added report section - Endpoint Backup Results Summary
# Added report section - Endpoint Backup Job Status
# Added report section - Running Endpoint Backup Jobs
# Added report section - Endpoint Backup Jobs/Sessions with Warnings or Failures
# Added report section - Successful Endpoint Backup Jobs/Sessions
#
# Version 1.4.1 - SM
# Fixed issue with summary counts
# Version 1.4 - SM
# Misc minor tweaks/cleanup
# Added variable for report width
# Added variable for email subject
# Added ability to show/hide all report sections
# Added Protected/Unprotected VM Count to Summary
# Added per object details for sessions w/no details
# Added proxy host name to Proxy Details
# Added repository host name to Repository Details
# Added section showing successful sessions
# Added ability to view only last session per job
# Added Cluster field for protected/unprotected VMs
# Added catch for cifs repositories greater than 4TB as erroneous data is returned
# Added % Complete for Running Jobs
# Added ability to exclude multiple (vCenter) folders from Missing and Successful Backups section
# Added ability to exclude multiple (vCenter) datacenters from Missing and Successful Backups section
# Tweaked license info for better reporting across different date formats
#
# Version 1.3 - SM
# Now supports VBR v8
# For VBR v7, use report version 1.2
# Added more flexible options to save and launch file
#
# Version 1.2 - SM
# Added option to show VMs Successfully backed up
#
# Version 1.1.4 - SM
# Misc tweaks/bug fixes
# Reconfigured HTML a bit to help with certain email clients
# Added cell coloring to highlight status
# Added $rptTitle variable to hold report title
# Added ability to send report via email as attachment
#
# Version 1.1.3 - SM
# Added Details to Sessions with Warnings or Failures
#
# Version 1.1.2 - SM
# Minor tweaks/updates
# Added Veeam version info to header
#
# Version 1.1.1 - Shawn Masterson
# Based on vPowerCLI v6 Army Report (v1.1) by Thomas McConnell
# http://www.vpowercli.co.uk/2012/01/23/vpowercli-v6-army-report/
# http://pastebin.com/6p3LrWt7
#
# Tweaked HTML header (color, title)
#
# Changed report width to 1024px
#
# Moved hard-coded path to exe/dll files to user declared variables ($veeamExePath/$veeamDllPath)
#
# Adjusted sorting on all objects
#
# Modified info group/counts
#   Modified - Total Jobs = Job Runs
#   Added - Read (GB)
#   Added - Transferred (GB)
#   Modified - Warning = Warnings
#   Modified - Failed = Failures
#   Added - Failed (last session)
#   Added - Running (currently running sessions)
#
# Modified job lines
#   Renamed Header - Sessions with Warnings or Failures
#   Fixed Write (GB) - Broke with v7
#
# Added support license renewal
#   Credit - Gavin Townsend  http://www.theagreeablecow.com/2012/09/sysadmin-modular-reporting-samreports.html
#   Original  Credit - Arne Fokkema  http://ict-freak.nl/2011/12/29/powershell-veeam-br-get-total-days-before-the-license-expires/
#
# Modified Proxy section
#   Removed Read/Write/Util - Broke in v7 - Workaround unknown
#
# Modified Services section
#   Added - $runningSvc variable to toggle displaying services that are running
#   Added - Ability to hide section if no results returned (all services are running)
#   Added - Scans proxies and repositories as well as the VBR server for services
#
# Added VMs Not Backed Up section
#   Credit - Tom Sightler - http://sightunseen.org/blog/?p=1
#   http://www.sightunseen.org/files/vm_backup_status_dev.ps1
#
# Modified $reportMode
#   Added ability to run with any number of hours (8,12,72 etc)
#   Added bits to allow for zero sessions (semi-gracefully)
#
# Added Running Jobs section
#   Added ability to toggle displaying running jobs
#
# Added catch to ensure running v7 or greater
#
#
# Version 1.1
# Added job lines as per a request on the website
#
# Version 1.0
# Clean up for release
#
# Version 0.9
# More cmdlet rewrite to improve perfomace, credit to @SethBartlett
# for practically writing the Get-vPCRepoInfo
#
# Version 0.8
# Added Read/Write stats for proxies at requests of @bsousapt
# Performance improvement of proxy tear down due to rewrite of cmdlet
# Replaced 2 other functions
# Added Warning counter, .00 to all storage returns and fetch credentials for
# remote WinLocal repos
#
# Version 0.7
# Added Utilisation(Get-vPCDailyProxyUsage) and Modes 24, 48, Weekly, and Monthly
# Minor performance tweaks
#endregion

#region Connect


# Connect to VBR server
Write-Host "Connecting ..."
$OpenConnection = (Get-VBRServerSession).Server
If ($OpenConnection -ne $vbrServer){
  Disconnect-VBRServer
  Try {
    Connect-VBRServer -server $vbrServer -ErrorAction Stop
  } Catch {
    Write-Host "Unable to connect to VBR server - $vbrServer" -ForegroundColor Red
    exit
  }
}
#endregion

#region NonUser-Variables
# Get all Backup/Backup Copy/Replica Jobs
$allJobs = @()
If ($showSummaryBk + $showJobsBk + $showFileJobsBk + $showAllSessBk + $showAllTasksBk + $showRunningBk +
  $showRunningTasksBk + $showWarnFailBk + $showTaskWFBk + $showSuccessBk + $showTaskSuccessBk +
  $showSummaryRp + $showJobsRp + $showAllSessRp + $showAllTasksRp + $showRunningRp +
  $showRunningTasksRp + $showWarnFailRp + $showTaskWFRp + $showSuccessRp + $showTaskSuccessRp +
  $showSummaryBc + $showJobsBc + $showAllSessBc + $showAllTasksBc + $showIdleBc +
  $showPendingTasksBc + $showRunningBc + $showRunningTasksBc + $showWarnFailBc +
  $showTaskWFBc + $showSuccessBc + $showTaskSuccessBc) {
  $allJobs = Get-VBRJob -WarningAction SilentlyContinue
}

#Other version where FileBackup is just added to normal backup job sessions.
#$allJobsBk = @($allJobs | Where-Object {$_.JobType -eq "Backup" -or $_.JobType -eq"NasBackup" })
# Get all Backup Jobs
$allJobsBk = @($allJobs | Where-Object {$_.JobType -eq "Backup"})
# Get all File Backup Jobs
$allFileJobsBk = @($allJobs | Where-Object {$_.JobType -eq "NasBackup"})
# Get all Replication Jobs
$allJobsRp = @($allJobs | Where-Object {$_.JobType -eq "Replica"})
# Get all Backup Copy Jobs
$allJobsBc = @($allJobs | Where-Object {$_.JobType -eq "BackupSync" -or $_.JobType -eq "SimpleBackupCopyPolicy"})
# Get all Tape Jobs
$allJobsTp = @()
If ($showSummaryTp + $showJobsTp + $showAllSessTp + $showAllTasksTp +
  $showWaitingTp + $showIdleTp + $showPendingTasksTp + $showRunningTp + $showRunningTasksTp +
  $showWarnFailTp + $showTaskWFTp + $showSuccessTp + $showTaskSuccessTp) {
  $allJobsTp = @(Get-VBRTapeJob)
}
# Get all Agent Backup Jobs
$allJobsEp = @()
If ($showSummaryEp + $showJobsEp + $showAllSessEp + $showRunningEp +
  $showWarnFailEp + $showSuccessEp) {
  $allJobsEp = @(Get-VBRComputerBackupJob)
}
# Get all SureBackup Jobs
$allJobsSb = @()
If ($showSummarySb + $showJobsSb + $showAllSessSb + $showAllTasksSb +
  $showRunningSb + $showRunningTasksSb + $showWarnFailSb + $showTaskWFSb +
  $showSuccessSb + $showTaskSuccessSb) {
  $allJobsSb = @(Get-VSBJob)
}

# Get all Backup/Backup Copy/Replica Sessions
$allSess = @()
If ($allJobs) {
  $allSess = Get-VBRBackupSession
}
# Get all File / NAS Backup Sessions
$allFileSess = @()
If ($allFileJobs) {
  $allFileSess = Get-VBRNASBackupSession -Name *
}


# Get all Restore Sessions
$allSessResto = @()
If ($showRestoRunVM + $showRestoreVM) {
  $allSessResto = Get-VBRRestoreSession
}
# Get all Tape Backup Sessions
$allSessTp = @()
If ($allJobsTp) {
  Foreach ($tpJob in $allJobsTp){
    $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
    $allSessTp += $tpSessions
  }
}
# Get all Agent Backup Sessions
$allSessEp = @()
If ($allJobsEp) {
  $allSessEp = Get-VBRComputerBackupJobSession
}
# Get all SureBackup Sessions
$allSessSb = @()
If ($allJobsSb) {
  $allSessSb = Get-VSBSession
}

# Get all Backups
$jobBackups = @()
If ($showBackupSizeBk + $showBackupSizeBc + $showBackupSizeEp) {
  $jobBackups = Get-VBRBackup
}
# Get Backup Job Backups
$backupsBk = @($jobBackups | Where-Object {$_.JobType -eq "Backup"})
# Get Backup Copy Job Backups
$backupsBc = @($jobBackups | Where-Object {$_.JobType -eq "BackupSync"})
# Get Agent Backup Job Backups
$backupsEp = @($jobBackups | Where-Object {$_.JobType -eq "EndpointBackup"})

# Get all Media Pools
$mediaPools = Get-VBRTapeMediaPool
# Get all Media Vaults
Try {
    $mediaVaults = Get-VBRTapeVault
} Catch {
    Write-Host "Tape possibly not licensed."
}
# Get all Tapes
$mediaTapes = Get-VBRTapeMedium
# Get all Tape Libraries
$mediaLibs = Get-VBRTapeLibrary
# Get all Tape Drives
$mediaDrives = Get-VBRTapeDrive

# Get Configuration Backup Info
$configBackup = Get-VBRConfigurationBackupJob
# Get VBR Server object
$vbrServerObj = Get-VBRLocalhost
# Get all Proxies
$proxyList = Get-VBRViProxy
# Get all Repositories
$repoList = Get-VBRBackupRepository
$repoListSo = Get-VBRBackupRepository -ScaleOut
# Get all Tape Servers
$tapesrvList = Get-VBRTapeServer

# Convert mode (timeframe) to hours
If ($reportMode -eq "Monthly") {
  $HourstoCheck = 720
} Elseif ($reportMode -eq "Weekly") {
  $HourstoCheck = 168
} Else {
  $HourstoCheck = $reportMode
}

# Gather all Backup Sessions within timeframe
$sessListBk = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($null -ne $backupJob -and $backupJob -ne "") {
  $allJobsBkTmp = @()
  $sessListBkTmp = @()
  $backupsBkTmp = @()
  Foreach ($bkJob in $backupJob) {
    $allJobsBkTmp += $allJobsBk | Where-Object {$_.Name -like $bkJob}
    $sessListBkTmp += $sessListBk | Where-Object {$_.JobName -like $bkJob}
    $backupsBkTmp += $backupsBk | Where-Object {$_.JobName -like $bkJob}
  }
  $allJobsBk = $allJobsBkTmp | Sort-Object Id -Unique
  $sessListBk = $sessListBkTmp | Sort-Object Id -Unique
  $backupsBk = $backupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
  $tempSessListBk = $sessListBk
  $sessListBk = @()
  Foreach($job in $allJobsBk) {
    $sessListBk += $tempSessListBk | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0

$sessListBk | ForEach-Object {$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBk | ForEach-Object {$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Success"})
$warningSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBk = @($sessListBk | Where-Object {$_.Result -eq "Failed"})
$runningSessionsBk = @($sessListBk | Where-Object {$_.State -eq "Working"})
$failedSessionsBk = @($sessListBk | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# File Backup Session Section Start

$fileSessListBk = @($allFileSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Backup"})
If ($null -ne $backupJob -and $backupJob -ne "") {
  $allFileJobsBkTmp = @()
  $fileSessListBkTmp = @()
  $fileBackupsBkTmp = @()
  Foreach ($bkJob in $backupJob) {
    $allFileJobsBkTmp += $allFileJobsBk | Where-Object {$_.Name -like $bkJob}
    $fileSessListBkTmp += $fileSessListBk | Where-Object {$_.JobName -like $bkJob}
    $fileBackupsBkTmp += $fileBackupsBk | Where-Object {$_.JobName -like $bkJob}
  }
  $allFileJobsBk = $allFileJobsBkTmp | Sort-Object Id -Unique
  $fileSessListBk = $fileSessListBkTmp | Sort-Object Id -Unique
  $fileBackupsBk = $fileBackupsBkTmp | Sort-Object Id -Unique
}
If ($onlyLastBk) {
  $tempFileSessListBk = $fileSessListBk
  $fileSessListBk = @()
  Foreach($job in $allFileJobsBk) {
    $fileSessListBk += $tempFileSessListBk | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Session information
$totalXferBk = 0
$totalReadBk = 0

$fileSessListBk | ForEach-Object {$totalXferBk += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$fileSessListBk | ForEach-Object {$totalReadBk += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successFileSessionsBk = @($fileSessListBk | Where-Object {$_.Result -eq "Success"})
$warningFileSessionsBk = @($fileSessListBk | Where-Object {$_.Result -eq "Warning"})
$failsFileSessionsBk = @($fileSessListBk | Where-Object {$_.Result -eq "Failed"})
$runningFileSessionsBk = @($fileSessListBk | Where-Object {$_.State -eq "Working"})
$failedFileSessionsBk = @($fileSessListBk | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})



# End File Backup Session Section End

# Gather VM Restore Sessions within timeframe
$sessListResto = @($allSessResto | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or !($_.IsCompleted)})
# Get VM Restore Session information
$completeResto = @($sessListResto | Where-Object {$_.IsCompleted})
$runningResto = @($sessListResto | Where-Object {!($_.IsCompleted)})

# Gather all Replication Sessions within timeframe
$sessListRp = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working") -and $_.JobType -eq "Replica"})
If ($null -ne $replicaJob -and $replicaJob -ne "") {
  $allJobsRpTmp = @()
  $sessListRpTmp = @()
  Foreach ($rpJob in $replicaJob) {
    $allJobsRpTmp += $allJobsRp | Where-Object {$_.Name -like $rpJob}
    $sessListRpTmp += $sessListRp | Where-Object {$_.JobName -like $rpJob}
  }
  $allJobsRp = $allJobsRpTmp | Sort-Object Id -Unique
  $sessListRp = $sessListRpTmp | Sort-Object Id -Unique
}
If ($onlyLastRp) {
  $tempSessListRp = $sessListRp
  $sessListRp = @()
  Foreach($job in $allJobsRp) {
    $sessListRp += $tempSessListRp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Replication Session information
$totalXferRp = 0
$totalReadRp = 0
$sessListRp | ForEach-Object {$totalXferRp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListRp | ForEach-Object {$totalReadRp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$successSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Success"})
$warningSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsRp = @($sessListRp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsRp = @($sessListRp | Where-Object {$_.State -eq "Working"})
$failedSessionsRp = @($sessListRp | Where-Object {($_.Result -eq "Failed") -and ($_.WillBeRetried -ne "True")})

# Gather all Backup Copy Sessions within timeframe
$sessListBc = @($allSess | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle") -and ($_.JobType -eq "BackupSync" -or $_.JobType -eq "SimpleBackupCopyWorker")})
If ($null -ne $bcopyJob -and $bcopyJob -ne "") {
  $allJobsBcTmp = @()
  $sessListBcTmp = @()
  $backupsBcTmp = @()
  Foreach ($bcJob in $bcopyJob) {
    $allJobsBcTmp += $allJobsBc | Where-Object {$_.Name -like $bcJob}
    $sessListBcTmp += $sessListBc | Where-Object {$_.JobName -like $bcJob}
    $backupsBcTmp += $backupsBc | Where-Object {$_.JobName -like $bcJob}
  }
  $allJobsBc = $allJobsBcTmp | Sort-Object Id -Unique
  $sessListBc = $sessListBcTmp | Sort-Object Id -Unique
  $backupsBc = $backupsBcTmp | Sort-Object Id -Unique
}
If ($onlyLastBc) {
  $tempSessListBc = $sessListBc
  $sessListBc = @()
  Foreach($job in $allJobsBc) {
    $sessListBc += $tempSessListBc | Where-Object {$_.Jobname -eq $job.name -and $_.BaseProgress -eq 100} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Backup Copy Session information
$totalXferBc = 0
$totalReadBc = 0
$sessListBc | ForEach-Object {$totalXferBc += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListBc | ForEach-Object {$totalReadBc += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Idle"})
$successSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Success"})
$warningSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Warning"})
$failsSessionsBc = @($sessListBc | Where-Object {$_.Result -eq "Failed"})
$workingSessionsBc = @($sessListBc | Where-Object {$_.State -eq "Working"})

# Gather all Tape Backup Sessions within timeframe
$sessListTp = @($allSessTp | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
If ($null -ne $tapeJob -and $tapeJob -ne "") {
  $allJobsTpTmp = @()
  $sessListTpTmp = @()
  Foreach ($tpJob in $tapeJob) {
    $allJobsTpTmp += $allJobsTp | Where-Object {$_.Name -like $tpJob}
    $sessListTpTmp += $sessListTp | Where-Object {$_.JobName -like $tpJob}
  }
  $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
  $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
}
If ($onlyLastTp) {
  $tempSessListTp = $sessListTp
  $sessListTp = @()
  Foreach($job in $allJobsTp) {
    $sessListTp += $tempSessListTp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Tape Backup Session information
$totalXferTp = 0
$totalReadTp = 0
$sessListTp | ForEach-Object {$totalXferTp += $([Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2))}
$sessListTp | ForEach-Object {$totalReadTp += $([Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2))}
$idleSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Idle"})
$successSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Success"})
$warningSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Failed"})
$workingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Working"})
$waitingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "WaitingTape"})

# Gather all Agent Backup Sessions within timeframe
$sessListEp = $allSessEp | Where-Object {($_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -eq "Working")}
If ($null -ne $epbJob -and $epbJob -ne "") {
  $allJobsEpTmp = @()
  $sessListEpTmp = @()
  $backupsEpTmp = @()
  Foreach ($eJob in $epbJob) {
    $allJobsEpTmp += $allJobsEp | Where-Object {$_.Name -like $eJob}
    $backupsEpTmp += $backupsEp | Where-Object {$_.JobName -like $eJob}
  }
  Foreach ($job in $allJobsEpTmp) {
    $sessListEpTmp += $sessListEp | Where-Object {$_.JobId -eq $job.Id}
  }
  $allJobsEp = $allJobsEpTmp | Sort-Object Id -Unique
  $sessListEp = $sessListEpTmp | Sort-Object Id -Unique
  $backupsEp = $backupsEpTmp | Sort-Object Id -Unique
}
If ($onlyLastEp) {
  $tempSessListEp = $sessListEp
  $sessListEp = @()
  Foreach($job in $allJobsEp) {
    $sessListEp += $tempSessListEp | Where-Object {$_.JobId -eq $job.Id} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get Agent Backup Session information
$successSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Success"})
$warningSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Warning"})
$failsSessionsEp = @($sessListEp | Where-Object {$_.Result -eq "Failed"})
$runningSessionsEp = @($sessListEp | Where-Object {$_.State -eq "Working"})

# Gather all SureBackup Sessions within timeframe
$sessListSb = @($allSessSb | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -ne "Stopped"})
If ($null -ne $surebJob -and $surebJob -ne "") {
  $allJobsSbTmp = @()
  $sessListSbTmp = @()
  Foreach ($SbJob in $surebJob) {
    $allJobsSbTmp += $allJobsSb | Where-Object {$_.Name -like $SbJob}
    $sessListSbTmp += $sessListSb | Where-Object {$_.JobName -like $SbJob}
  }
  $allJobsSb = $allJobsSbTmp | Sort-Object Id -Unique
  $sessListSb = $sessListSbTmp | Sort-Object Id -Unique
}
If ($onlyLastSb) {
  $tempSessListSb = $sessListSb
  $sessListSb = @()
  Foreach($job in $allJobsSb) {
    $sessListSb += $tempSessListSb | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
  }
}
# Get SureBackup Session information
$successSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Success"})
$warningSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Warning"})
$failsSessionsSb = @($sessListSb | Where-Object {$_.Result -eq "Failed"})
$runningSessionsSb = @($sessListSb | Where-Object {$_.State -ne "Stopped"})

# Format Report Mode for header
If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
  $rptMode = "RPO: $reportMode Hrs"
} Else {
  $rptMode = "RPO: $reportMode"
}

# Toggle VBR Server name in report header
If ($showVBR) {
  $vbrName = "VBR Server - $vbrServer"
} Else {
  $vbrName = $null
}

# Append Report Mode to Email subject
If ($modeSubject) {
  If (($reportMode -ne "Weekly") -And ($reportMode -ne "Monthly")) {
    $emailSubject = "$emailSubject (Last $reportMode Hrs)"
  } Else {
    $emailSubject = "$emailSubject ($reportMode)"
  }
}

# Append VBR Server to Email subject
If ($vbrSubject) {
  $emailSubject = "$emailSubject - $vbrServer"
}

# Append Date and Time to Email subject
If ($dtSubject) {
  $emailSubject = "$emailSubject - $(Get-Date -format g)"
}
#endregion

#region Functions

Function Get-VBRProxyInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Proxy
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param ([PsObject]$inputObj)
      $ping = New-Object system.net.networkinformation.ping
      $isIP = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
      If ($inputObj.Host.Name -match $isIP) {
        $IPv4 = $inputObj.Host.Name
      } Else {
        $DNS = [Net.DNS]::GetHostEntry("$($inputObj.Host.Name)")
        $IPv4 = ($DNS.get_AddressList() | Where-Object {$_.AddressFamily -eq "InterNetwork"} | Select-Object -First 1).IPAddressToString
      }
      $pinginfo = $ping.send("$($IPv4)")
      If ($pinginfo.Status -eq "Success") {
        $hostAlive = "Alive"
        $response = $pinginfo.RoundtripTime
      } Else {
        $hostAlive = "Dead"
        $response = $null
      }
      If ($inputObj.IsDisabled) {
        $enabled = "False"
      } Else {
        $enabled = "True"
      }
      $tMode = switch ($inputObj.Options.TransportMode) {
        "Auto" {"Automatic"}
        "San" {"Direct SAN"}
        "HotAdd" {"Hot Add"}
        "Nbd" {"Network"}
        default {"Unknown"}
      }
      $vPCFuncObject = New-Object PSObject -Property @{
        ProxyName = $inputObj.Name
        RealName = $inputObj.Host.Name.ToLower()
        Disabled = $inputObj.IsDisabled
        pType = $inputObj.ChassisType
        Status  = $hostAlive
        IP = $IPv4
        Response = $response
        Enabled = $enabled
        maxtasks = $inputObj.Options.MaxTasksCount
        tMode = $tMode
      }
      Return $vPCFuncObject
    }
  }
  Process {
    Foreach ($p in $Proxy) {
      $outputObj = Build-Object $p
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }
}

Function Get-VBRRepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $repohost, $path, $free, $total, $maxtasks, $rtype)
      $repoObj = New-Object -TypeName PSObject -Property @{
        Target = $name
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        StorageBackup = [Math]::Round([Decimal]$rBackupsize/1GB,2)
        StorageOther = [Math]::Round([Decimal]($total-$rBackupsize-$free)/1GB-0.5,2)
        MaxTasks = $maxtasks
        rType = $rtype
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($r in $Repository) {
      # Refresh Repository Size Info
      [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
      $rType = switch ($r.Type) {
        "WinLocal" {"Windows Local"}
        "LinuxLocal" {"Linux Local"}
        "LinuxHardened" {"Hardened"}
        "CifsShare" {"CIFS Share"}
        "DataDomain" {"Data Domain"}
        "ExaGrid" {"ExaGrid"}
        "HPStoreOnce" {"HP StoreOnce"}
        "Nfs" {"NFS Direct"}
        default {"Unknown"}
      }
      $outputObj = Build-Object $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $r.Options.MaxTaskCount $rType
    }
    $outputAry += $outputObj
  }
  End {
    $outputAry
  }
}

Function Get-VBRSORepoInfo {
  [CmdletBinding()]
  param (
    [Parameter(Position=0, ValueFromPipeline=$true)]
    [PSObject[]]$Repository
  )
  Begin {
    $outputAry = @()
    Function Build-Object {param($name, $rname, $repohost, $path, $free, $total, $maxtasks, $rtype, $capenabled)
      $repoObj = New-Object -TypeName PSObject -Property @{
        SoTarget = $name
        Target = $rname
        RepoHost = $repohost
        Storepath = $path
        StorageFree = [Math]::Round([Decimal]$free/1GB,2)
        StorageTotal = [Math]::Round([Decimal]$total/1GB,2)
        FreePercentage = [Math]::Round(($free/$total)*100)
        MaxTasks = $maxtasks
        rType = $rtype
        capEnabled = $capenabled
      }
      Return $repoObj
    }
  }
  Process {
    Foreach ($rs in $Repository) {
      ForEach ($rp in $rs.Extent) {
        $r = $rp.Repository
        # Refresh Repository Size Info
        [Veeam.Backup.Core.CBackupRepositoryEx]::SyncSpaceInfoToDb($r, $true)
		$rBackupSize = [Veeam.Backup.Core.CBackupRepository]::GetRepositoryBackupsSize($r.Id.Guid)
        $rType = switch ($r.Type) {
          "WinLocal" {"Windows Local"}
          "LinuxLocal" {"Linux Local"}
          "LinuxHardened" {"Hardened"}
          "CifsShare" {"CIFS Share"}
          "DataDomain" {"Data Domain"}
          "ExaGrid" {"ExaGrid"}
          "HPStoreOnce" {"HPE StoreOnce"}
          "Nfs" {"NFS Direct"}
          "SanSnapshotOnly" {"SAN Snapshot"}
          "Cloud" {"VCSP Cloud"}
          default {"Unknown"}
        }
		if ($rtype -eq "SAN Snapshot" -or $rtype -eq "VCSP Cloud") {$maxTaskCount="N/A"}
		else {$maxTaskCount=$r.Options.MaxTaskCount}
        $outputObj = Build-Object $rs.Name $r.Name $($r.GetHost()).Name.ToLower() $r.Path $r.GetContainer().CachedFreeSpace.InBytes $r.GetContainer().CachedTotalSpace.InBytes $maxTaskCount $rType $rBackupSize
        $outputAry += $outputObj
      }
    <# #Added for capacity tier begin ToDo
    if($rs.CapacityExtent.Repository.Name.Length -gt 0) {
        $ce = $rs.CapacityExtent
        $outputObj = Build-Object $rs.Name $ce.Repository.Name $ce.Repository.ServicePoint $ce.Repository.AmazonS3Folder
        $outputAry += $outputObj
    }
    #Added for capacity tier end #>
    }
  }
  End {
    $outputAry
  }
}

function Get-RepoPermission {
  $outputAry = @()
  $repoEPPerms = $script:repoList | get-vbreppermission
  $repoEPPermsSo = $script:repoListSo | get-vbreppermission
  ForEach ($repo in $repoEPPerms) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = (Get-VBRBackupRepository | Where-Object {$_.Id -eq $repo.RepositoryId}).Name
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  ForEach ($repo in $repoEPPermsSo) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Name = "[SO] $((Get-VBRBackupRepository -ScaleOut | Where-Object {$_.Id -eq $repo.RepositoryId}).Name)"
      "Permission Type" = $repo.PermissionType
      Users = $repo.Users | Out-String
      "Encryption Enabled" = $repo.IsEncryptionEnabled
    }
    $outputAry += $objoutput
  }
  $outputAry
}

Function Get-VBRReplicaTarget {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipeline=$true)]
    [PSObject[]]$InputObj
  )
  BEGIN {
    $outputAry = @()
    $dsAry = @()
    If (($null -ne $Name) -and ($null -ne $InputObj)) {
      $InputObj = Get-VBRJob -Name $Name
    }
  }
  PROCESS {
    Foreach ($obj in $InputObj) {
      If (($dsAry -contains $obj.ViReplicaTargetOptions.DatastoreName) -eq $false) {
        $esxi = $obj.GetTargetHost()
        $dtstr =  $esxi | Find-VBRViDatastore -Name $obj.ViReplicaTargetOptions.DatastoreName
        $objoutput = New-Object -TypeName PSObject -Property @{
          Target = $esxi.Name
          Datastore = $obj.ViReplicaTargetOptions.DatastoreName
          StorageFree = [Math]::Round([Decimal]$dtstr.FreeSpace/1GB,2)
          StorageTotal = [Math]::Round([Decimal]$dtstr.Capacity/1GB,2)
          FreePercentage = [Math]::Round(($dtstr.FreeSpace/$dtstr.Capacity)*100)
        }
        $dsAry = $dsAry + $obj.ViReplicaTargetOptions.DatastoreName
        $outputAry = $outputAry + $objoutput
      } Else {
        return
      }
    }
  }
  END {
    $outputAry | Select-Object Target, Datastore, StorageFree, StorageTotal, FreePercentage
  }
}

Function Get-VeeamVersion {
  Try {
    $veeamCore = Get-Item -Path $veeamCorePath
    $VeeamVersion = [single]($veeamCore.VersionInfo.ProductVersion).substring(0,4)
    $productVersion=[string]$veeamCore.VersionInfo.ProductVersion
    $productHotfix=[string]$veeamCore.VersionInfo.Comments
    Write-Host "Found Veeam Version: $productVersion $productHotfix"
    $objectVersion = New-Object -TypeName PSObject -Property @{
          VeeamVersion = $VeeamVersion
          productVersion = $productVersion
          productHotfix = $productHotfix
    }

    Return $objectVersion
  } Catch {
    Write-Error "Unable to Locate Veeam Core, check path - $veeamCorePath" -ForegroundColor Red
    exit
  }
}

Function Get-VeeamSupportDate {
    # Query for license info
    $licenseInfo = Get-VBRInstalledLicense

    $type = $licenseinfo.Type

    switch ( $type ) {
        'Perpetual' {
            $date = $licenseInfo.SupportExpirationDate
        }
        'Evaluation' {
            # No expiration
            $date = Get-Date
        }
        'Subscription' {
            $date = $licenseInfo.ExpirationDate
        }
        'Rental' {
            $date = $licenseInfo.ExpirationDate
        }
        'NFR' {
            $date = $licenseInfo.ExpirationDate
        }

    }

    [PSCustomObject]@{
       LicType    = $type
       ExpDate    = $date.ToShortDateString()
       DaysRemain = ($date - (Get-Date)).Days
    }
}

Function Get-VeeamWinServer {
  $vservers=@{}
  $outputAry = @()
  $vservers.add($($script:vbrServerObj.Name),"VBRServer")
  Foreach ($srv in $script:proxyList) {
    If (!$vservers.ContainsKey($srv.Host.Name)) {
      $vservers.Add($srv.Host.Name,"ProxyServer")
    }
  }
  Foreach ($srv in $script:repoList) {
    If ($srv.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($srv.gethost().Name)) {
      $vservers.Add($srv.gethost().Name,"RepoServer")
    }
  }
  Foreach ($rs in $script:repoListSo) {
    ForEach ($rp in $rs.Extent) {
      $r = $rp.Repository
      $rName = $($r.GetHost()).Name
      If ($r.Type -ne "LinuxLocal" -and !$vservers.ContainsKey($rName)) {
        $vservers.Add($rName,"RepoSoServer")
      }
    }
  }
  Foreach ($srv in $script:tapesrvList) {
    If (!$vservers.ContainsKey($srv.Name)) {
      $vservers.Add($srv.Name,"TapeServer")
    }
  }
  $vservers = $vservers.GetEnumerator() | Sort-Object Name
  Foreach ($vserver in $vservers) {
    $outputAry += $vserver.Name
  }
  return $outputAry
}

Function Get-VeeamService {
  param (
    [PSObject]$inputObj
  )
  $outputAry = @()
  Foreach ($obj in $InputObj) {
    $output = @()
    Try {
      $output = Get-Service -computername $obj -Name "*Veeam*" -exclude "SQLAgent*" |
        Select-Object @{Name="Server Name"; Expression = {$obj.ToLower()}}, @{Name="Service Name"; Expression = {$_.DisplayName}}, Status
    } Catch {
      $output = New-Object PSObject -Property @{
        "Server Name" = $obj.ToLower()
        "Service Name" = "Unable to connect"
        Status = "Unknown"
      }
    }
    $outputAry += $output
  }
  $outputAry
}

Function Get-VMsBackupStatus {
  $outputary = @()
  # Convert exclusion list to simple regular expression
  $excludevms_regex = ('(?i)^(' + (($script:excludeVMs | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $excludefolder_regex = ('(?i)^(' + (($script:excludeFolder | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  #ToDo: exclude VMs based on tags
  $excludedc_regex = ('(?i)^(' + (($script:excludeDC | ForEach-Object {[regex]::escape($_)}) -join "|") + ')$') -replace "\\\*", ".*"
  $vms=@{}
  # Build a hash table of all VMs.  Key is either Job Object Id (for any VM ever in a Veeam job) or vCenter ID+MoRef
  # Assume unprotected (!), and populate Cluster, DataCenter, and Name fields for hash key value
  Find-VBRViEntity |
    Where-Object {$_.Type -eq "Vm" -and $_.VmFolderName -notmatch $excludefolder_regex} |
    Where-Object {$_.Name -notmatch $excludevms_regex} |
    Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
    ForEach-Object {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.Path.Split("\")[2], $_.Name, "1/11/1911", "1/11/1911","", $_.VmFolderName))}

  If (!$script:excludeTemp) {
    Find-VBRViEntity -VMsandTemplates |
      Where-Object {$_.Type -eq "Vm" -and $_.IsTemplate -eq "True" -and $_.VmFolderName -notmatch $excludefolder_regex} |
      Where-Object {$_.Name -notmatch $excludevms_regex} |
      Where-Object {$_.Path.Split("\")[1] -notmatch $excludedc_regex} |
      ForEach-Object {$vms.Add(($_.FindObject().Id, $_.Id -ne $null)[0], @("!", $_.Path.Split("\")[0], $_.Path.Split("\")[1], $_.VmHostName, "[template] $($_.Name)", "1/11/1911", "1/11/1911","", $_.VmFolderName))}
  }
  # Find all backup task sessions that have ended in the last x hours
  $vbrtasksessions = (Get-VBRBackupSession |
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
    Get-VBRTaskSession | Where-Object {$_.Status -notmatch "InProgress|Pending"}
  # Compare VM list to session list and update found VMs status
  If ($vbrtasksessions) {
    ForEach ($vmtask in $vbrtasksessions) {
      If ($vms.ContainsKey($vmtask.Info.ObjectId)) {
        If ((Get-Date $vmtask.Progress.StartTimeLocal) -ge (Get-Date $vms[$vmtask.Info.ObjectId][5])) {
          If ($vmtask.Status -eq "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=""
          } ElseIf ($vms[$vmtask.Info.ObjectId][0] -ne "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=($vmtask.GetDetails()).Replace("<br />","ZZbrZZ")
          }
        } ElseIf ($vms[$vmtask.Info.ObjectId][0] -match "Warning|Failed" -and $vmtask.Status -eq "Success") {
            $vms[$vmtask.Info.ObjectId][0]=$vmtask.Status
            $vms[$vmtask.Info.ObjectId][5]=$vmtask.Progress.StartTimeLocal
            $vms[$vmtask.Info.ObjectId][6]=$vmtask.Progress.StopTimeLocal
            $vms[$vmtask.Info.ObjectId][7]=""
        }
      }
    }
  }
  Foreach ($vm in $vms.GetEnumerator()) {
    $objoutput = New-Object -TypeName PSObject -Property @{
      Status = $vm.Value[0]
      Name = $vm.Value[4]
      vCenter = $vm.Value[1]
      Datacenter = $vm.Value[2]
      Cluster = $vm.Value[3]
      StartTime = $vm.Value[5]
      StopTime = $vm.Value[6]
      Details = $vm.Value[7]
      Folder = $vm.Value[8]
    }
    $outputAry += $objoutput
  }
  $outputAry
}

function Get-Duration {
  param ($ts)
  $days = ""
  If ($ts.Days -gt 0) {
    $days = "{0}:" -f $ts.Days
  }
  "{0}{1}:{2,2:D2}:{3,2:D2}" -f $days,$ts.Hours,$ts.Minutes,$ts.Seconds
}

function Get-BackupSize {
  param ($backups)
  $outputObj = @()
  Foreach ($backup in $backups) {
    $backupSize = 0
    $dataSize = 0
    $files = $backup.GetAllStorages()
    Foreach ($file in $Files) {
      $backupSize += [math]::Round([long]$file.Stats.BackupSize/1GB, 2)
      $dataSize += [math]::Round([long]$file.Stats.DataSize/1GB, 2)
    }
    $repo = If ($($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name) {
              $($script:repoList | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
            } Else {
              $($script:repoListSo | Where-Object {$_.Id -eq $backup.RepositoryId}).Name
            }
    $vbrMasterHash = @{
      JobName = $backup.JobName
      VMCount = $backup.VmCount
      Repo = $repo
      DataSize = $dataSize
      BackupSize = $backupSize
    }
    $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
    $outputObj += $vbrMasterObj
  }
  $outputObj
}

Function Get-MultiJob {
  $outputAry = @()
  $vmMultiJobs = (Get-VBRBackupSession |
    Where-Object {($_.JobType -eq "Backup") -and ($_.EndTime -ge (Get-Date).addhours(-$script:HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$script:HourstoCheck) -or $_.State -eq "Working")}) |
    Get-VBRTaskSession | Select-Object Name, @{Name="VMID"; Expression = {$_.Info.ObjectId}}, JobName -Unique | Group-Object Name, VMID | Where-Object {$_.Count -gt 1} | Select-Object -ExpandProperty Group
  ForEach ($vm in $vmMultiJobs) {
    $objID = $vm.VMID
    $viEntity = Find-VBRViEntity -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
    If ($null -ne $viEntity) {
      $objoutput = New-Object -TypeName PSObject -Property @{
        Name = $vm.Name
        vCenter = $viEntity.Path.Split("\")[0]
        Datacenter = $viEntity.Path.Split("\")[1]
        Cluster = $viEntity.Path.Split("\")[2]
        Folder = $viEntity.VMFolderName
        JobName = $vm.JobName
      }
      $outputAry += $objoutput
    } Else { #assume Template
      $viEntity = Find-VBRViEntity -VMsAndTemplates -name $vm.Name | Where-Object {$_.FindObject().Id -eq $objID}
      If ($null -ne $viEntity) {
        $objoutput = New-Object -TypeName PSObject -Property @{
          Name = "[template] " + $vm.Name
          vCenter = $viEntity.Path.Split("\")[0]
          Datacenter = $viEntity.Path.Split("\")[1]
          Cluster = $viEntity.VmHostName
          Folder = $viEntity.VMFolderName
          JobName = $vm.JobName
        }
      }
      If ($objoutput) {
        $outputAry += $objoutput
      }
    }
  }
  $outputAry
}
#endregion

#region Report
# Get Veeam Version
$objectVersion = Get-VeeamVersion

If ($objectVersion.VeeamVersion -lt 11.0) {
  Write-Host "Script requires VBR v11.0 or higher" -ForegroundColor Red
  exit
}


Write-Host "Generating reports, please be patient ..."

# HTML Stuff
$headerObj = @"
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <title>$rptTitle</title>
            <style type="text/css">
              body {font-family: Tahoma; background-color: #ffffff;}
              table {font-family: Tahoma; width: $($rptWidth)%; font-size: 12px; border-collapse: collapse; margin-left: auto; margin-right: auto;}
              table tr:nth-child(odd) td {background: $oddColor;}
              th {background-color: #e2e2e2; border: 1px solid #a7a9ac;border-bottom: none;}
              td {background-color: #ffffff; border: 1px solid #a7a9ac;padding: 2px 3px 2px 3px;}
            </style>
    </head>
"@

$bodyTop = @"
    <body>
          <table>
              <tr>
                  <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 10px;vertical-align: bottom;text-align: left;padding: 2px 0px 0px 5px;"></td>
                  <td style="width: 50%;height: 14px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 2px 5px 0px 0px;">Report generated on $(Get-Date -format g)</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 24px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 15px;">$rptTitle</td>
                  <td style="width: 50%;height: 24px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">$vbrName</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 0px 5px;"></td>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 0px 0px;">VBR v$($objectVersion.productVersion)</td>
              </tr>
              <tr>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: left;padding: 0px 0px 2px 5px;">$rptMode</td>
                  <td style="width: 50%;height: 12px;border: none;background-color: ZZhdbgZZ;color: White;font-size: 12px;vertical-align: bottom;text-align: right;padding: 0px 5px 2px 0px;">MVR v$MVRversion</td>
              </tr>
          </table>
"@

$subHead01 = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #f3f4f4;color: #626365;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01suc = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #00b050;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01war = @"
<table>
                 <tr>
                    <td style="height: 35px;background-color: #ffd96c;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01err = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #FB9895;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead01inf = @"
<table>
                <tr>
                    <td style="height: 35px;background-color: #3399FF;color: #ffffff;font-size: 16px;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;">
"@

$subHead02 = @"
</td>
                </tr>
             </table>
"@

$HTMLbreak = @"
<table>
                <tr>
                    <td style="height: 10px;background-color: #626365;padding: 5px 0 0 15px;border-top: 5px solid white;border-bottom: none;"></td>
						    </tr>
            </table>
"@

$footerObj = @"
            <table>
                <tr>
                    <td style="height: 15px;background-color: #ffffff;border: none;color: #626365;font-size: 10px;text-align:center;">My Veeam Report developed by <a href="http://blog.smasterson.com" target="_blank">http://blog.smasterson.com</a> and modified for V11 by <a href="http://horstmann.in" target="_blank">http://horstmann.in</a></td>
                </tr>
            </table>
    </body>
</html>
"@

#region Get VM Backup Status
$vmStatus = @()
If ($showSummaryProtect + $showUnprotectedVMs + $showUnprotectedVMsInfo + $showProtectedVMs) {
  $vmStatus = Get-VMsBackupStatus
}
# VMs Missing Backups
$missingVMs = @($vmStatus | Where-Object {$_.Status -match "!|Failed"})
ForEach ($VM in $missingVMs) {
  If ($VM.Status -eq "!") {
    $VM.Details = "No Backup Task has completed"
    $VM.StartTime = ""
    $VM.StopTime = ""
  }
}
# VMs Successfuly Backed Up
$successVMs = @($vmStatus | Where-Object {$_.Status -eq "Success"})
# VMs Backed Up w/Warning
$warnVMs = @($vmStatus | Where-Object {$_.Status -eq "Warning"})
#endregion

#region Get VM Backup Protection Summary
$bodySummaryProtect = $null
$sumprotectHead = $subHead01
If ($showSummaryProtect) {
  If (@($successVMs).Count -ge 1) {
    $percentProt = 1
    $sumprotectHead = $subHead01suc
  }
  If (@($warnVMs).Count -ge 1) {
    $percentWarn = "*"
    $sumprotectHead = $subHead01war
  } Else {
    $percentWarn = ""
  }
  If (@($missingVMs).Count -ge 1) {
    $percentProt = (@($warnVMs).Count + @($successVMs).Count) / (@($warnVMs).Count + @($successVMs).Count + @($missingVMs).Count)
    If ($showUnprotectedVMsInfo) {
      $sumprotectHead = $subHead01inf
    } Else {
      $sumprotectHead = $subHead01err
    }
  }
  $vbrMasterHash = @{
    WarningVM = @($warnVMs).Count
    ProtectedVM = @($successVMs).Count
    UnprotectedVM = @($missingVMs).Count
    PercentProt = "{0:P2}{1}" -f $percentProt,$percentWarn

  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  $summaryProtect =  $vbrMasterObj | Select-Object @{Name="% Protected"; Expression = {$_.PercentProt}},
    @{Name="Fully Protected VMs"; Expression = {$_.ProtectedVM}},
    @{Name="Protected VMs w/Warnings"; Expression = {$_.WarningVM}},
    @{Name="Unprotected VMs"; Expression = {$_.UnprotectedVM}}
  $bodySummaryProtect = $summaryProtect | ConvertTo-HTML -Fragment
  $bodySummaryProtect = $sumprotectHead + "VM Backup Protection Summary" + $subHead02 + $bodySummaryProtect
}
#endregion

#region Get VMs Missing Backups
$bodyMissing = $null
If ($showUnprotectedVMs -Or $showUnprotectedVMsInfo) {
  If ($missingVMs.count -gt 0) {

    If ($showUnprotectedVMsInfo) {
      $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-HTML -Fragment
      $bodyMissing = $subHead01inf + "Unprotected VMs within RPO" + $subHead02 + $missingVMs
    } Else{
      $missingVMs = $missingVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
        @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-HTML -Fragment
      $bodyMissing = $subHead01err + "VMs with No Successful Backups within RPO" + $subHead02 + $missingVMs
    }
  }
}
#endregion

# Get VMs Backed Up w/Warnings
$bodyWarning = $null
If ($showProtectedVMs) {
  If ($warnVMs.Count -gt 0) {
    $warnVMs = $warnVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}}, Details | ConvertTo-HTML -Fragment
    $bodyWarning = $subHead01war + "VMs with only Backups with Warnings within RPO" + $subHead02 + $warnVMs
  }
}

# Get VMs Successfuly Backed Up
$bodySuccess = $null
If ($showProtectedVMs) {
  If ($successVMs.Count -gt 0) {
    $successVMs = $successVMs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Last Start Time"; Expression = {$_.StartTime}}, @{Name="Last End Time"; Expression = {$_.StopTime}} | ConvertTo-HTML -Fragment
    $bodySuccess = $subHead01suc + "VMs with Successful Backups within RPO" + $subHead02 + $successVMs
  }
}

# Get VMs Backed Up by Multiple Jobs
$bodyMultiJobs = $null
If ($showMultiJobs) {
  $multiJobs = @(Get-MultiJob)
  If ($multiJobs.Count -gt 0) {
    $bodyMultiJobs = $multiJobs | Sort-Object vCenter, Datacenter, Cluster, Name | Select-Object Name, vCenter, Datacenter, Cluster, Folder,
      @{Name="Job Name"; Expression = {$_.JobName}} | ConvertTo-HTML -Fragment
    $bodyMultiJobs = $subHead01err + "VMs Backed Up by Multiple Jobs within RPO" + $subHead02 + $bodyMultiJobs
  }
}

# Get Backup Summary Info
$bodySummaryBk = $null
If ($showSummaryBk) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsBk).Count
    "Sessions" = If ($sessListBk) {@($sessListBk).Count} Else {0}
    "Read" = $totalReadBk
    "Transferred" = $totalXferBk
    "Successful" = @($successSessionsBk).Count
    "Warning" = @($warningSessionsBk).Count
    "Fails" = @($failsSessionsBk).Count
    "Running" = @($runningSessionsBk).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBk) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBk =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryBk = $arrSummaryBk | ConvertTo-HTML -Fragment
  If ($arrSummaryBk.Failed -gt 0) {
      $summaryBkHead = $subHead01err
  } ElseIf ($arrSummaryBk.Warnings -gt 0) {
      $summaryBkHead = $subHead01war
  } ElseIf ($arrSummaryBk.Successful -gt 0) {
      $summaryBkHead = $subHead01suc
  } Else {
      $summaryBkHead = $subHead01
  }
  $bodySummaryBk = $summaryBkHead + "Backup Results Summary" + $subHead02 + $bodySummaryBk
}

# Get Backup Job Status
$bodyJobsBk = $null
If ($showJobsBk) {
  If ($allJobsBk.count -gt 0) {
    $bodyJobsBk = @()
    Foreach($bkJob in $allJobsBk) {
      $bodyJobsBk += $bkJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($bkJob.IsRunning) {
            $currentSess = $runningSessionsBk | Where-Object {$_.JobName -eq $bkJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name) {
            $($repoList | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          } Else {
            $($repoListSo | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          }
        }},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinuous) {"<Continuous>"}
		  ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name + "]"}
		  Else {$_.ScheduleOptions.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){"Unknown"}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBk = $bodyJobsBk | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsBk = $subHead01 + "Backup Job Status" + $subHead02 + $bodyJobsBk
  }
}

# Get Backup Job Status Begin
$bodyFileJobsBk = $null
If ($showFileJobsBk) {
  If ($allFileJobsBk.count -gt 0) {
    $bodyFileJobsBk = @()
    Foreach($bkJob in $allFileJobsBk) {
      $bodyFileJobsBk += $bkJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($bkJob.IsRunning) {
            $currentSess = $runningSessionsBk | Where-Object {$_.JobName -eq $bkJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name) {
            $($repoList | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          } Else {
            $($repoListSo | Where-Object {$_.Id -eq $BkJob.Info.TargetRepositoryId}).Name
          }
        }},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinuous) {"<Continuous>"}
		  ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $bkJob.Info.ParentScheduleId}).Name + "]"}
		  Else {$_.ScheduleOptions.NextRun}
        }},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){"Unknown"}Else{$_.Info.LatestStatus}}}
    }
    $bodyFileJobsBk = $bodyFileJobsBk | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
    $bodyFileJobsBk = $subHead01 + "File Backup Job Status" + $subHead02 + $bodyFileJobsBk
  }
}
# Get File Backup Job Status End

# Get Backup Job Size Begin
$bodyJobSizeBk = $null
If ($showBackupSizeBk) {
  If ($backupsBk.count -gt 0) {
    $bodyJobSizeBk = Get-BackupSize -backups $backupsBk | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBk = $subHead01 + "Backup Job Size" + $subHead02 + $bodyJobSizeBk
  }
}
# Get Backup Job Size End

# Get File Backup Job Size Begin
$bodyFileJobSizeBk = $null
If ($showFileBackupSizeBk) {
  If ($fileBackupsBk.count -gt 0) {
    $bodyFileJobSizeBk = Get-BackupSize -backups $fileBackupsBk | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyFileJobSizeBk = $subHead01 + "File Backup Job Size" + $subHead02 + $bodyFileJobSizeBk
  }
}
# Get Backup Job Size End


# Get all Backup Sessions
$bodyAllSessBk = $null
If ($showAllSessBk) {
  If ($sessListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk  | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    } Else {
      $arrAllSessBk = $sessListBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBk = $arrAllSessBk | ConvertTo-HTML -Fragment
      If ($arrAllSessBk.Result -match "Failed") {
        $allSessBkHead = $subHead01err
      } ElseIf ($arrAllSessBk.Result -match "Warning") {
        $allSessBkHead = $subHead01war
      } ElseIf ($arrAllSessBk.Result -match "Success") {
        $allSessBkHead = $subHead01suc
      } Else {
        $allSessBkHead = $subHead01
      }
      $bodyAllSessBk = $allSessBkHead + "Backup Sessions" + $subHead02 + $bodyAllSessBk
    }
  }
}

# Get Running Backup Jobs
$bodyRunningBk = $null
If ($showRunningBk) {
  If ($runningSessionsBk.count -gt 0) {
    $bodyRunningBk = $runningSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBk = $subHead01 + "Running Backup Jobs" + $subHead02 + $bodyRunningBk
  }
}

# Get Backup Sessions with Warnings or Failures
$bodySessWFBk = $null
If ($showWarnFailBk) {
  $sessWF = @($warningSessionsBk + $failsSessionsBk)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBk) {
      $headerWF = "Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedBk) {
      $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    } Else {
      $arrSessWFBk = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBk = $arrSessWFBk | ConvertTo-HTML -Fragment
      If ($arrSessWFBk.Result -match "Failed") {
        $sessWFBkHead = $subHead01err
      } ElseIf ($arrSessWFBk.Result -match "Warning") {
        $sessWFBkHead = $subHead01war
      } ElseIf ($arrSessWFBk.Result -match "Success") {
        $sessWFBkHead = $subHead01suc
      } Else {
        $sessWFBkHead = $subHead01
      }
      $bodySessWFBk = $sessWFBkHead + $headerWF + $subHead02 + $bodySessWFBk
    }
  }
}

# Get Successful Backup Sessions
$bodySessSuccBk = $null
If ($showSuccessBk) {
  If ($successSessionsBk.count -gt 0) {
    If ($onlyLastBk) {
      $headerSucc = "Successful Backup Jobs"
    } Else {
      $headerSucc = "Successful Backup Sessions"
    }
    If ($showDetailedBk) {
      $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    } Else {
      $bodySessSuccBk = $successSessionsBk | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBk = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBk
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Tasks from Sessions within time frame
$taskListBk = @()
$taskListBk += $sessListBk | Get-VBRTaskSession
$successTasksBk = @($taskListBk | Where-Object {$_.Status -eq "Success"})
$wfTasksBk = @($taskListBk | Where-Object {$_.Status -match "Warning|Failed"})
$runningTasksBk = @()
$runningTasksBk += $runningSessionsBk | Get-VBRTaskSession | Where-Object {$_.Status -match "Pending|InProgress"}

# Get all Backup Tasks
$bodyAllTasksBk = $null
If ($showAllTasksBk) {
  If ($taskListBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    } Else {
      $arrAllTasksBk = $taskListBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBk = $arrAllTasksBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBk.Status -match "Failed") {
        $allTasksBkHead = $subHead01err
      } ElseIf ($arrAllTasksBk.Status -match "Warning") {
        $allTasksBkHead = $subHead01war
      } ElseIf ($arrAllTasksBk.Status -match "Success") {
        $allTasksBkHead = $subHead01suc
      } Else {
        $allTasksBkHead = $subHead01
      }
      $bodyAllTasksBk = $allTasksBkHead + "Backup Tasks" + $subHead02 + $bodyAllTasksBk
    }
  }
}

# Get Running Backup Tasks
$bodyTasksRunningBk = $null
If ($showRunningTasksBk) {
  If ($runningTasksBk.count -gt 0) {
    $bodyTasksRunningBk = $runningTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBk = $subHead01 + "Running Backup Tasks" + $subHead02 + $bodyTasksRunningBk
  }
}

# Get Backup Tasks with Warnings or Failures
$bodyTaskWFBk = $null
If ($showTaskWFBk) {
  If ($wfTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    } Else {
      $arrTaskWFBk = $wfTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBk = $arrTaskWFBk | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBk.Status -match "Failed") {
        $taskWFBkHead = $subHead01err
      } ElseIf ($arrTaskWFBk.Status -match "Warning") {
        $taskWFBkHead = $subHead01war
      } ElseIf ($arrTaskWFBk.Status -match "Success") {
        $taskWFBkHead = $subHead01suc
      } Else {
        $taskWFBkHead = $subHead01
      }
      $bodyTaskWFBk = $taskWFBkHead + "Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBk
    }
  }
}

# Get Successful Backup Tasks
$bodyTaskSuccBk = $null
If ($showTaskSuccessBk) {
  If ($successTasksBk.count -gt 0) {
    If ($showDetailedBk) {
      $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    } Else {
      $bodyTaskSuccBk = $successTasksBk | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBk = $subHead01suc + "Successful Backup Tasks" + $subHead02 + $bodyTaskSuccBk
    }
  }
}

# Get Running VM Restore Sessions
$bodyRestoRunVM = $null
If ($showRestoRunVM) {
  If ($($runningResto).count -gt 0) {
    $bodyRestoRunVM = $runningResto | Sort-Object CreationTime | Select-Object @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}}, @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}},
      @{Name="Reason"; Expression = {$_.Info.Reason}} | ConvertTo-HTML -Fragment
    $bodyRestoRunVM = $subHead01 + "Running VM Restore Sessions" + $subHead02 + $bodyRestoRunVM
  }
}

# Get Completed VM Restore Sessions
$bodyRestoreVM = $null
If ($showRestoreVM) {
  If ($($completeResto).count -gt 0) {
    $arrRestoreVM = $completeResto | Sort-Object CreationTime | Select-Object @{Name="VM Name"; Expression = {$_.Info.VmDisplayName}},
      @{Name="Restore Type"; Expression = {$_.JobTypeString}},
      @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
      @{Name="Initiator"; Expression = {$_.Info.Initiator.Name}}, @{Name="Reason"; Expression = {$_.Info.Reason}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyRestoreVM = $arrRestoreVM | ConvertTo-HTML -Fragment
    If ($arrRestoreVM.Result -match "Failed") {
      $restoreVMHead = $subHead01err
    } ElseIf ($arrRestoreVM.Result -match "Warning") {
      $restoreVMHead = $subHead01war
    } ElseIf ($arrRestoreVM.Result -match "Success") {
      $restoreVMHead = $subHead01suc
    } Else {
      $restoreVMHead = $subHead01
    }
    $bodyRestoreVM = $restoreVMHead + "Completed VM Restore Sessions" + $subHead02 + $bodyRestoreVM
  }
}

# Get Replication Summary Info
$bodySummaryRp = $null
If ($showSummaryRp) {
  $vbrMasterHash = @{
    "Failed" = @($failedSessionsRp).Count
    "Sessions" = If ($sessListRp) {@($sessListRp).Count} Else {0}
    "Read" = $totalReadRp
    "Transferred" = $totalXferRp
    "Successful" = @($successSessionsRp).Count
    "Warning" = @($warningSessionsRp).Count
    "Fails" = @($failsSessionsRp).Count
    "Running" = @($runningSessionsRp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastRp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryRp =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}},
    @{Name="Failed"; Expression = {$_.Failed}}
  $bodySummaryRp = $arrSummaryRp | ConvertTo-HTML -Fragment
  If ($arrSummaryRp.Failed -gt 0) {
      $summaryRpHead = $subHead01err
  } ElseIf ($arrSummaryRp.Warnings -gt 0) {
      $summaryRpHead = $subHead01war
  } ElseIf ($arrSummaryRp.Successful -gt 0) {
      $summaryRpHead = $subHead01suc
  } Else {
      $summaryRpHead = $subHead01
  }
  $bodySummaryRp = $summaryRpHead + "Replication Results Summary" + $subHead02 + $bodySummaryRp
}

# Get Replication Job Status
$bodyJobsRp = $null
If ($showJobsRp) {
  If ($allJobsRp.count -gt 0) {
    $bodyJobsRp = @()
    Foreach($rpJob in $allJobsRp) {
      $bodyJobsRp += $rpJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($rpJob.IsRunning) {
            $currentSess = $runningSessionsRp | Where-Object {$_.JobName -eq $rpJob.Name}
            $csessPercent = $currentSess.Progress.Percents
            $csessSpeed = [Math]::Round($currentSess.Info.Progress.AvgSpeed/1MB,2)
            $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
            $cStatus
          } Else {
            "Stopped"
          }
         }},
        @{Name="Target"; Expression = {$(Get-VBRServer | Where-Object {$_.Id -eq $rpJob.Info.TargetHostId}).Name}},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where-Object {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name) {$($repoList | Where-Object {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where-Object {$_.Id -eq $rpJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinuous) {"<Continuous>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $rpJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsRp = $bodyJobsRp | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsRp = $subHead01 + "Replication Job Status" + $subHead02 + $bodyJobsRp
  }
}

# Get Replication Sessions
$bodyAllSessRp = $null
If ($showAllSessRp) {
  If ($sessListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    } Else {
      $arrAllSessRp = $sessListRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessRp = $arrAllSessRp | ConvertTo-HTML -Fragment
      If ($arrAllSessRp.Result -match "Failed") {
        $allSessRpHead = $subHead01err
      } ElseIf ($arrAllSessRp.Result -match "Warning") {
        $allSessRpHead = $subHead01war
      } ElseIf ($arrAllSessRp.Result -match "Success") {
        $allSessRpHead = $subHead01suc
      } Else {
        $allSessRpHead = $subHead01
      }
      $bodyAllSessRp = $allSessRpHead + "Replication Sessions" + $subHead02 + $bodyAllSessRp
    }
  }
}

# Get Running Replication Jobs
$bodyRunningRp = $null
If ($showRunningRp) {
  If ($runningSessionsRp.count -gt 0) {
    $bodyRunningRp = $runningSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningRp = $subHead01 + "Running Replication Jobs" + $subHead02 + $bodyRunningRp
  }
}

# Get Replication Sessions with Warnings or Failures
$bodySessWFRp = $null
If ($showWarnFailRp) {
  $sessWF = @($warningSessionsRp + $failsSessionsRp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastRp) {
      $headerWF = "Replication Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Replication Sessions with Warnings or Failures"
    }
    If ($showDetailedRp) {
      $arrSessWFRp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    } Else {
      $arrSessWFRp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFRp = $arrSessWFRp | ConvertTo-HTML -Fragment
      If ($arrSessWFRp.Result -match "Failed") {
        $sessWFRpHead = $subHead01err
      } ElseIf ($arrSessWFRp.Result -match "Warning") {
        $sessWFRpHead = $subHead01war
      } ElseIf ($arrSessWFRp.Result -match "Success") {
        $sessWFRpHead = $subHead01suc
      } Else {
        $sessWFRpHead = $subHead01
      }
      $bodySessWFRp = $sessWFRpHead + $headerWF + $subHead02 + $bodySessWFRp
    }
  }
}

# Get Successful Replication Sessions
$bodySessSuccRp = $null
If ($showSuccessRp) {
  If ($successSessionsRp.count -gt 0) {
    If ($onlyLastRp) {
      $headerSucc = "Successful Replication Jobs"
    } Else {
      $headerSucc = "Successful Replication Sessions"
    }
    If ($showDetailedRp) {
      $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    } Else {
      $bodySessSuccRp = $successSessionsRp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccRp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccRp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Replication Tasks from Sessions within time frame
$taskListRp = @()
$taskListRp += $sessListRp | Get-VBRTaskSession
$successTasksRp = @($taskListRp | Where-Object {$_.Status -eq "Success"})
$wfTasksRp = @($taskListRp | Where-Object {$_.Status -match "Warning|Failed"})
$runningTasksRp = @()
$runningTasksRp += $runningSessionsRp | Get-VBRTaskSession | Where-Object {$_.Status -match "Pending|InProgress"}

# Get Replication Tasks
$bodyAllTasksRp = $null
If ($showAllTasksRp) {
  If ($taskListRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrAllTasksRp = $taskListRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    } Else {
      $arrAllTasksRp = $taskListRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksRp = $arrAllTasksRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksRp.Status -match "Failed") {
        $allTasksRpHead = $subHead01err
      } ElseIf ($arrAllTasksRp.Status -match "Warning") {
        $allTasksRpHead = $subHead01war
      } ElseIf ($arrAllTasksRp.Status -match "Success") {
        $allTasksRpHead = $subHead01suc
      } Else {
        $allTasksRpHead = $subHead01
      }
      $bodyAllTasksRp = $allTasksRpHead + "Replication Tasks" + $subHead02 + $bodyAllTasksRp
    }
  }
}

# Get Running Replication Tasks
$bodyTasksRunningRp = $null
If ($showRunningTasksRp) {
  If ($runningTasksRp.count -gt 0) {
    $bodyTasksRunningRp = $runningTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningRp = $subHead01 + "Running Replication Tasks" + $subHead02 + $bodyTasksRunningRp
  }
}

# Get Replication Tasks with Warnings or Failures
$bodyTaskWFRp = $null
If ($showTaskWFRp) {
  If ($wfTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $arrTaskWFRp = $wfTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    } Else {
      $arrTaskWFRp = $wfTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFRp = $arrTaskWFRp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFRp.Status -match "Failed") {
        $taskWFRpHead = $subHead01err
      } ElseIf ($arrTaskWFRp.Status -match "Warning") {
        $taskWFRpHead = $subHead01war
      } ElseIf ($arrTaskWFRp.Status -match "Success") {
        $taskWFRpHead = $subHead01suc
      } Else {
        $taskWFRpHead = $subHead01
      }
      $bodyTaskWFRp = $taskWFRpHead + "Replication Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFRp
    }
  }
}

# Get Successful Replication Tasks
$bodyTaskSuccRp = $null
If ($showTaskSuccessRp) {
  If ($successTasksRp.count -gt 0) {
    If ($showDetailedRp) {
      $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    } Else {
      $bodyTaskSuccRp = $successTasksRp | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccRp = $subHead01suc + "Successful Replication Tasks" + $subHead02 + $bodyTaskSuccRp
    }
  }
}

# Get Backup Copy Summary Info
$bodySummaryBc = $null
If ($showSummaryBc) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListBc) {@($sessListBc).Count} Else {0}
    "Read" = $totalReadBc
    "Transferred" = $totalXferBc
    "Successful" = @($successSessionsBc).Count
    "Warning" = @($warningSessionsBc).Count
    "Fails" = @($failsSessionsBc).Count
    "Working" = @($workingSessionsBc).Count
    "Idle" = @($idleSessionsBc).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastBc) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryBc =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryBc = $arrSummaryBc | ConvertTo-HTML -Fragment
  If ($arrSummaryBc.Failures -gt 0) {
      $summaryBcHead = $subHead01err
  } ElseIf ($arrSummaryBc.Warnings -gt 0) {
      $summaryBcHead = $subHead01war
  } ElseIf ($arrSummaryBc.Successful -gt 0) {
      $summaryBcHead = $subHead01suc
  } Else {
      $summaryBcHead = $subHead01
  }
  $bodySummaryBc = $summaryBcHead + "Backup Copy Results Summary" + $subHead02 + $bodySummaryBc
}

# Get Backup Copy Job Status
$bodyJobsBc = $null
If ($showJobsBc) {
  If ($allJobsBc.count -gt 0) {
    $bodyJobsBc = @()
    Foreach($BcJob in $allJobsBc) {
      $bodyJobsBc += $BcJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.Info.IsScheduleEnabled}},
        @{Name="Type"; Expression = {$_.TypeToString}},
        @{Name="Status"; Expression = {
          If ($BcJob.IsRunning) {
            $currentSess = $BcJob.FindLastSession()
            If ($currentSess.State -eq "Working") {
              $csessPercent = $currentSess.Progress.Percents
              $csessSpeed = [Math]::Round($currentSess.Progress.AvgSpeed/1MB,2)
              $cStatus = "$($csessPercent)% completed at $($csessSpeed) MB/s"
              $cStatus
            } Else {
              $currentSess.State
            }
          } Else {
            "Stopped"
          }
        }},
        @{Name="Target Repo"; Expression = {
          If ($($repoList | Where-Object {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name) {$($repoList | Where-Object {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}
          Else {$($repoListSo | Where-Object {$_.Id -eq $BcJob.Info.TargetRepositoryId}).Name}}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.Options.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinious) {"<Continious>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $BcJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.Info.LatestStatus -eq "None"){""}Else{$_.Info.LatestStatus}}}
    }
    $bodyJobsBc = $bodyJobsBc | Sort-Object "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsBc = $subHead01 + "Backup Copy Job Status" + $subHead02 + $bodyJobsBc
  }
}

# Get Backup Copy Job Size
$bodyJobSizeBc = $null
If ($showBackupSizeBc) {
  If ($backupsBc.count -gt 0) {
    $bodyJobSizeBc = Get-BackupSize -backups $backupsBc | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeBc = $subHead01 + "Backup Copy Job Size" + $subHead02 + $bodyJobSizeBc
  }
}

# Get All Backup Copy Sessions
$bodyAllSessBc = $null
If ($showAllSessBc) {
  If ($sessListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    } Else {
      $arrAllSessBc = $sessListBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessBc = $arrAllSessBc | ConvertTo-HTML -Fragment
      If ($arrAllSessBc.Result -match "Failed") {
        $allSessBcHead = $subHead01err
      } ElseIf ($arrAllSessBc.Result -match "Warning") {
        $allSessBcHead = $subHead01war
      } ElseIf ($arrAllSessBc.Result -match "Success") {
        $allSessBcHead = $subHead01suc
      } Else {
        $allSessBcHead = $subHead01
      }
      $bodyAllSessBc = $allSessBcHead + "Backup Copy Sessions" + $subHead02 + $bodyAllSessBc
    }
  }
}

# Get Idle Backup Copy Sessions
$bodySessIdleBc = $null
If ($showIdleBc) {
  If ($idleSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerIdle = "Idle Backup Copy Jobs"
    } Else {
      $headerIdle = "Idle Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    } Else {
      $bodySessIdleBc = $idleSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleBc = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleBc
    }
  }
}

# Get Working Backup Copy Jobs
$bodyRunningBc = $null
If ($showRunningBc) {
  If ($workingSessionsBc.count -gt 0) {
    $bodyRunningBc = $workingSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningBc = $subHead01 + "Working Backup Copy Sessions" + $subHead02 + $bodyRunningBc
  }
}

# Get Backup Copy Sessions with Warnings or Failures
$bodySessWFBc = $null
If ($showWarnFailBc) {
  $sessWF = @($warningSessionsBc + $failsSessionsBc)
  If ($sessWF.count -gt 0) {
    If ($onlyLastBc) {
      $headerWF = "Backup Copy Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Backup Copy Sessions with Warnings or Failures"
    }
    If ($showDetailedBc) {
      $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    } Else {
      $arrSessWFBc = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFBc = $arrSessWFBc | ConvertTo-HTML -Fragment
      If ($arrSessWFBc.Result -match "Failed") {
        $sessWFBcHead = $subHead01err
      } ElseIf ($arrSessWFBc.Result -match "Warning") {
        $sessWFBcHead = $subHead01war
      } ElseIf ($arrSessWFBc.Result -match "Success") {
        $sessWFBcHead = $subHead01suc
      } Else {
        $sessWFBcHead = $subHead01
      }
      $bodySessWFBc = $sessWFBcHead + $headerWF + $subHead02 + $bodySessWFBc
    }
  }
}

# Get Successful Backup Copy Sessions
$bodySessSuccBc = $null
If ($showSuccessBc) {
  If ($successSessionsBc.count -gt 0) {
    If ($onlyLastBc) {
      $headerSucc = "Successful Backup Copy Jobs"
    } Else {
      $headerSucc = "Successful Backup Copy Sessions"
    }
    If ($showDetailedBc) {
      $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Dedupe"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetDedupeX(),1)) +"x"}}},
        @{Name="Compression"; Expression = {
          If ($_.Progress.ReadSize -eq 0) {0}
          Else {([string][Math]::Round($_.BackupStats.GetCompressX(),1)) +"x"}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    } Else {
      $bodySessSuccBc = $successSessionsBc | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccBc = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccBc
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Backup Copy Tasks from Sessions within time frame
$taskListBc = @()
$taskListBc += $sessListBc | Get-VBRTaskSession
$successTasksBc = @($taskListBc | Where-Object {$_.Status -eq "Success"})
$wfTasksBc = @($taskListBc | Where-Object {$_.Status -match "Warning|Failed"})
$pendingTasksBc = @($taskListBc | Where-Object {$_.Status -eq "Pending"})
$runningTasksBc = @($taskListBc | Where-Object {$_.Status -eq "InProgress"})

# Get All Backup Copy Tasks
$bodyAllTasksBc = $null
If ($showAllTasksBc) {
  If ($taskListBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrAllTasksBc = $taskListBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    } Else {
      $arrAllTasksBc = $taskListBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksBc = $arrAllTasksBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksBc.Status -match "Failed") {
        $allTasksBcHead = $subHead01err
      } ElseIf ($arrAllTasksBc.Status -match "Warning") {
        $allTasksBcHead = $subHead01war
      } ElseIf ($arrAllTasksBc.Status -match "Success") {
        $allTasksBcHead = $subHead01suc
      } Else {
        $allTasksBcHead = $subHead01
      }
      $bodyAllTasksBc = $allTasksBcHead + "Backup Copy Tasks" + $subHead02 + $bodyAllTasksBc
    }
  }
}

# Get Pending Backup Copy Tasks
$bodyTasksPendingBc = $null
If ($showPendingTasksBc) {
  If ($pendingTasksBc.count -gt 0) {
    $bodyTasksPendingBc = $pendingTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingBc = $subHead01 + "Pending Backup Copy Tasks" + $subHead02 + $bodyTasksPendingBc
  }
}

# Get Working Backup Copy Tasks
$bodyTasksRunningBc = $null
If ($showRunningTasksBc) {
  If ($runningTasksBc.count -gt 0) {
    $bodyTasksRunningBc = $runningTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningBc = $subHead01 + "Working Backup Copy Tasks" + $subHead02 + $bodyTasksRunningBc
  }
}

# Get Backup Copy Tasks with Warnings or Failures
$bodyTaskWFBc = $null
If ($showTaskWFBc) {
  If ($wfTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $arrTaskWFBc = $wfTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    } Else {
      $arrTaskWFBc = $wfTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFBc = $arrTaskWFBc | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFBc.Status -match "Failed") {
        $taskWFBcHead = $subHead01err
      } ElseIf ($arrTaskWFBc.Status -match "Warning") {
        $taskWFBcHead = $subHead01war
      } ElseIf ($arrTaskWFBc.Status -match "Success") {
        $taskWFBcHead = $subHead01suc
      } Else {
        $taskWFBcHead = $subHead01
      }
      $bodyTaskWFBc = $taskWFBcHead + "Backup Copy Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFBc
    }
  }
}

# Get Successful Backup Copy Tasks
$bodyTaskSuccBc = $null
If ($showTaskSuccessBc) {
  If ($successTasksBc.count -gt 0) {
    If ($showDetailedBc) {
      $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Processed (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedUsedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    } Else {
      $bodyTaskSuccBc = $successTasksBc | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccBc = $subHead01suc + "Successful Backup Copy Tasks" + $subHead02 + $bodyTaskSuccBc
    }
  }
}

# Get Tape Backup Summary Info
$bodySummaryTp = $null
If ($showSummaryTp) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListTp) {@($sessListTp).Count} Else {0}
    "Read" = $totalReadTp
    "Transferred" = $totalXferTp
    "Successful" = @($successSessionsTp).Count
    "Warning" = @($warningSessionsTp).Count
    "Fails" = @($failsSessionsTp).Count
    "Working" = @($workingSessionsTp).Count
    "Idle" = @($idleSessionsTp).Count
    "Waiting" = @($waitingSessionsTp).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastTp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryTp =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Read (GB)"; Expression = {$_.Read}}, @{Name="Transferred (GB)"; Expression = {$_.Transferred}},
    @{Name="Idle"; Expression = {$_.Idle}}, @{Name="Waiting"; Expression = {$_.Waiting}},
    @{Name="Working"; Expression = {$_.Working}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryTp = $arrSummaryTp | ConvertTo-HTML -Fragment
  If ($arrSummaryTp.Failures -gt 0) {
      $summaryTpHead = $subHead01err
  } ElseIf ($arrSummaryTp.Warnings -gt 0 -or $arrSummaryTp.Waiting -gt 0) {
      $summaryTpHead = $subHead01war
  } ElseIf ($arrSummaryTp.Successful -gt 0) {
      $summaryTpHead = $subHead01suc
  } Else {
      $summaryTpHead = $subHead01
  }
  $bodySummaryTp = $summaryTpHead + "Tape Backup Results Summary" + $subHead02 + $bodySummaryTp
}

# Get Tape Backup Job Status
$bodyJobsTp = $null
If ($showJobsTp) {
  If ($allJobsTp.count -gt 0) {
    $bodyJobsTp = @()
    Foreach($tpJob in $allJobsTp) {
      $bodyJobsTp += $tpJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Job Type"; Expression = {$_.Type}},@{Name="Media Pool"; Expression = {$_.Target}},
        @{Name="Status"; Expression = {$_.LastState}},
        @{Name="Next Run"; Expression = {
          If ($_.ScheduleOptions.Type -eq "AfterNewBackup") {"<Continuous>"}
          ElseIf ($_.ScheduleOptions.Type -eq "AfterJob") {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $tpJob.ScheduleOptions.JobId}).Name + "]"}
          ElseIf ($_.NextRun) {$_.NextRun}
          Else {"<not scheduled>"}}},
        @{Name="Last Result"; Expression = {If ($_.LastResult -eq "None"){""}Else{$_.LastResult}}}
    }
    $bodyJobsTp = $bodyJobsTp | Sort-Object "Next Run", "Job Name" | ConvertTo-HTML -Fragment
    $bodyJobsTp = $subHead01 + "Tape Backup Job Status" + $subHead02 + $bodyJobsTp
  }
}

# Get Tape Backup Sessions
$bodyAllSessTp = $null
If ($showAllSessTp) {
  If ($sessListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    } Else {
      $arrAllSessTp = $sessListTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Result
      $bodyAllSessTp = $arrAllSessTp | ConvertTo-HTML -Fragment
      If ($arrAllSessTp.Result -match "Failed") {
        $allSessTpHead = $subHead01err
      } ElseIf ($arrAllSessTp.Result -match "Warning" -or $arrAllSessTp.State -match "WaitingTape") {
        $allSessTpHead = $subHead01war
      } ElseIf ($arrAllSessTp.Result -match "Success") {
        $allSessTpHead = $subHead01suc
      } Else {
        $allSessTpHead = $subHead01
      }
      $bodyAllSessTp = $allSessTpHead + "Tape Backup Sessions" + $subHead02 + $bodyAllSessTp
    }

    # Due to issue with getting details on tape sessions, we may need to get session info again :-(
    If (($showWaitingTp -or $showIdleTp -or $showRunningTp -or $showWarnFailTp -or $showSuccessTp) -and $showDetailedTp) {
      # Get all Tape Backup Sessions
      $allSessTp = @()
      Foreach ($tpJob in $allJobsTp){
        $tpSessions = [veeam.backup.core.cbackupsession]::GetByJob($tpJob.id)
        $allSessTp += $tpSessions
      }
      # Gather all Tape Backup Sessions within timeframe
      $sessListTp = @($allSessTp | Where-Object {$_.EndTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.CreationTime -ge (Get-Date).AddHours(-$HourstoCheck) -or $_.State -match "Working|Idle"})
      If ($null -ne $tapeJob -and $tapeJob -ne "") {
        $allJobsTpTmp = @()
        $sessListTpTmp = @()
        Foreach ($tpJob in $tapeJob) {
          $allJobsTpTmp += $allJobsTp | Where-Object {$_.Name -like $tpJob}
          $sessListTpTmp += $sessListTp | Where-Object {$_.JobName -like $tpJob}
        }
        $allJobsTp = $allJobsTpTmp | Sort-Object Id -Unique
        $sessListTp = $sessListTpTmp | Sort-Object Id -Unique
      }
      If ($onlyLastTp) {
        $tempSessListTp = $sessListTp
        $sessListTp = @()
        Foreach($job in $allJobsTp) {
          $sessListTp += $tempSessListTp | Where-Object {$_.Jobname -eq $job.name} | Sort-Object EndTime -Descending | Select-Object -First 1
        }
      }
      # Get Tape Backup Session information
      $idleSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Idle"})
      $successSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Success"})
      $warningSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Warning"})
      $failsSessionsTp = @($sessListTp | Where-Object {$_.Result -eq "Failed"})
      $workingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "Working"})
      $waitingSessionsTp = @($sessListTp | Where-Object {$_.State -eq "WaitingTape"})
    }
  }
}

# Get Waiting Tape Backup Jobs
$bodyWaitingTp = $null
If ($showWaitingTp) {
  If ($waitingSessionsTp.count -gt 0) {
    $bodyWaitingTp = $waitingSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyWaitingTp = $subHead01war + "Waiting Tape Backup Sessions" + $subHead02 + $bodyWaitingTp
  }
}

# Get Idle Tape Backup Sessions
$bodySessIdleTp = $null
If ($showIdleTp) {
  If ($idleSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerIdle = "Idle Tape Backup Jobs"
    } Else {
      $headerIdle = "Idle Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    } Else {
      $bodySessIdleTp = $idleSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}} | ConvertTo-HTML -Fragment
      $bodySessIdleTp = $subHead01 + $headerIdle + $subHead02 + $bodySessIdleTp
    }
  }
}

# Get Working Tape Backup Jobs
$bodyRunningTp = $null
If ($showRunningTp) {
  If ($workingSessionsTp.count -gt 0) {
    $bodyRunningTp = $workingSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Progress.StartTimeLocal $(Get-Date))}},
      @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
      @{Name="Read (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.ReadSize/1GB, 2)}},
      @{Name="Transferred (GB)"; Expression = {[Math]::Round([Decimal]$_.Progress.TransferedSize/1GB, 2)}},
      @{Name="% Complete"; Expression = {$_.Progress.Percents}} | ConvertTo-HTML -Fragment
    $bodyRunningTp = $subHead01 + "Working Tape Backup Sessions" + $subHead02 + $bodyRunningTp
  }
}

# Get Tape Backup Sessions with Warnings or Failures
$bodySessWFTp = $null
If ($showWarnFailTp) {
  $sessWF = @($warningSessionsTp + $failsSessionsTp)
  If ($sessWF.count -gt 0) {
    If ($onlyLastTp) {
      $headerWF = "Tape Backup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "Tape Backup Sessions with Warnings or Failures"
    }
    If ($showDetailedTp) {
      $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    } Else {
      $arrSessWFTp = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}, Result
      $bodySessWFTp =  $arrSessWFTp | ConvertTo-HTML -Fragment
      If ($arrSessWFTp.Result -match "Failed") {
        $sessWFTpHead = $subHead01err
      } ElseIf ($arrSessWFTp.Result -match "Warning") {
        $sessWFTpHead = $subHead01war
      } ElseIf ($arrSessWFTp.Result -match "Success") {
        $sessWFTpHead = $subHead01suc
      } Else {
        $sessWFTpHead = $subHead01
      }
      $bodySessWFTp = $sessWFTpHead + $headerWF + $subHead02 + $bodySessWFTp
    }
  }
}

# Get Successful Tape Backup Sessions
$bodySessSuccTp = $null
If ($showSuccessTp) {
  If ($successSessionsTp.count -gt 0) {
    If ($onlyLastTp) {
      $headerSucc = "Successful Tape Backup Jobs"
    } Else {
      $headerSucc = "Successful Tape Backup Sessions"
    }
    If ($showDetailedTp) {
      $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Info.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Info.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Info.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Info.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result  | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    } Else {
      $bodySessSuccTp = $successSessionsTp | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {
          If ($_.GetDetails() -eq ""){$_ | Get-VBRTaskSession | ForEach-Object {If ($_.GetDetails()){$_.Name + ": " + ($_.GetDetails()).Replace("<br />","ZZbrZZ")}}}
          Else {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}},
        Result | ConvertTo-HTML -Fragment
      $bodySessSuccTp = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccTp
    }
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all Tape Backup Tasks from Sessions within time frame
$taskListTp = @()
$taskListTp += $sessListTp | Get-VBRTaskSession
$successTasksTp = @($taskListTp | Where-Object {$_.Status -eq "Success"})
$wfTasksTp = @($taskListTp | Where-Object {$_.Status -match "Warning|Failed"})
$pendingTasksTp = @($taskListTp | Where-Object {$_.Status -eq "Pending"})
$runningTasksTp = @($taskListTp | Where-Object {$_.Status -eq "InProgress"})

# Get Tape Backup Tasks
$bodyAllTasksTp = $null
If ($showAllTasksTp) {
  If ($taskListTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrAllTasksTp = $taskListTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    } Else {
      $arrAllTasksTp = $taskListTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Progress.StopTimeLocal}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyAllTasksTp = $arrAllTasksTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrAllTasksTp.Status -match "Failed") {
        $allTasksTpHead = $subHead01err
      } ElseIf ($arrAllTasksTp.Status -match "Warning") {
        $allTasksTpHead = $subHead01war
      } ElseIf ($arrAllTasksTp.Status -match "Success") {
        $allTasksTpHead = $subHead01suc
      } Else {
        $allTasksTpHead = $subHead01
      }
      $bodyAllTasksTp = $allTasksTpHead + "Tape Backup Tasks" + $subHead02 + $bodyAllTasksTp
    }
  }
}

# Get Pending Tape Backup Tasks
$bodyTasksPendingTp = $null
If ($showPendingTasksTp) {
  If ($pendingTasksTp.count -gt 0) {
    $bodyTasksPendingTp = $pendingTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksPendingTp = $subHead01 + "Pending Tape Backup Tasks" + $subHead02 + $bodyTasksPendingTp
  }
}

# Get Working Tape Backup Tasks
$bodyTasksRunningTp = $null
If ($showRunningTasksTp) {
  If ($runningTasksTp.count -gt 0) {
    $bodyTasksRunningTp = $runningTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Info.Progress.StartTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningTp = $subHead01 + "Working Tape Backup Tasks" + $subHead02 + $bodyTasksRunningTp
  }
}

# Get Tape Backup Tasks with Warnings or Failures
$bodyTaskWFTp = $null
If ($showTaskWFTp) {
  If ($wfTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $arrTaskWFTp = $wfTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    } Else {
      $arrTaskWFTp = $wfTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {$_.Progress.StopTimeLocal}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $_.Progress.Duration}},
        @{Name="Details"; Expression = {($_.GetDetails()).Replace("<br />","ZZbrZZ")}}, Status
      $bodyTaskWFTp = $arrTaskWFTp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      If ($arrTaskWFTp.Status -match "Failed") {
        $taskWFTpHead = $subHead01err
      } ElseIf ($arrTaskWFTp.Status -match "Warning") {
        $taskWFTpHead = $subHead01war
      } ElseIf ($arrTaskWFTp.Status -match "Success") {
        $taskWFTpHead = $subHead01suc
      } Else {
        $taskWFTpHead = $subHead01
      }
      $bodyTaskWFTp = $taskWFTpHead + "Tape Backup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFTp
    }
  }
}

# Get Successful Tape Backup Tasks
$bodyTaskSuccTp = $null
If ($showTaskSuccessTp) {
  If ($successTasksTp.count -gt 0) {
    If ($showDetailedTp) {
      $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        @{Name="Avg Speed (MB/s)"; Expression = {[Math]::Round($_.Progress.AvgSpeed/1MB,2)}},
        @{Name="Total (GB)"; Expression = {[Math]::Round($_.Progress.ProcessedSize/1GB,2)}},
        @{Name="Data Read (GB)"; Expression = {[Math]::Round($_.Progress.ReadSize/1GB,2)}},
        @{Name="Transferred (GB)"; Expression = {[Math]::Round($_.Progress.TransferedSize/1GB,2)}},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    } Else {
      $bodyTaskSuccTp = $successTasksTp | Select-Object @{Name="Name"; Expression = {$_.Name}},
        @{Name="Job Name"; Expression = {$_.JobSess.Name}},
        @{Name="Start Time"; Expression = {$_.Progress.StartTimeLocal}},
        @{Name="Stop Time"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {$_.Progress.StopTimeLocal}
        }},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.Progress.StopTimeLocal -eq "1/1/1900 12:00:00 AM") {"-"}
          Else {Get-Duration -ts $_.Progress.Duration}
        }},
        Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
      $bodyTaskSuccTp = $subHead01suc + "Successful Tape Backup Tasks" + $subHead02 + $bodyTaskSuccTp
    }
  }
}

# Get all Tapes
$bodyTapes = $null
If ($showTapes) {
  $expTapes = @($mediaTapes)
  If ($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select-Object Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort-Object Name | ConvertTo-HTML -Fragment
    $bodyTapes = $subHead01 + "All Tapes" + $subHead02 + $expTapes
  }
}

# Get all Tapes in each Custom/GFS Media Pool
$bodyTpPool = $null
If ($showTpMp) {
  ForEach ($mp in ($mediaPools | Where-Object {$_.Type -eq "Custom" -or $_.Type -eq "GFS"} | Sort-Object Name)) {
    $expTapes = @($mediaTapes | Where-Object {($_.MediaPoolId -eq $mp.Id)})
    If ($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select-Object Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpPool += $subHead01 + "All Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes in each Vault
$bodyTpVlt = $null
If ($showTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort-Object Name)) {
    $expTapes = @($mediaTapes | Where-Object {($_.Location.VaultId -eq $vlt.Id)})
    If ($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select-Object Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}},
      @{Name="Expiration Date"; Expression = {
          If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
            "Expired"
          } Else {
            $_.ExpirationDate
          }
      }} | Sort-Object Name | ConvertTo-HTML -Fragment
      $bodyTpVlt += $subHead01 + "All Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Expired Tapes
$bodyExpTp = $null
If ($showExpTp) {
  $expTapes = @($mediaTapes | Where-Object {($_.IsExpired -eq $True)})
  If ($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select-Object Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort-Object Name | ConvertTo-HTML -Fragment
    $bodyExpTp = $subHead01 + "All Expired Tapes" + $subHead02 + $expTapes
  }
}

# Get Expired Tapes in each Custom Media Pool
$bodyTpExpPool = $null
If ($showExpTpMp) {
  ForEach ($mp in ($mediaPools | Where-Object {$_.Type -eq "Custom"} | Sort-Object Name)) {
    $expTapes = @($mediaTapes | Where-Object {($_.MediaPoolId -eq $mp.Id -and $_.IsExpired -eq $True)})
    If ($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select-Object Name, Barcode,
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Location"; Expression = {
          switch ($_.Location) {
            "None" {"Offline"}
            "Slot" {
              $lId = $_.LibraryId
              $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
              [int]$slot = $_.SlotAddress + 1
              "{0} : {1} {2}" -f $lName,$_,$slot
            }
            "Drive" {
              $lId = $_.LibraryId
              $dId = $_.DriveId
              $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
              $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
              [int]$dNum = $_.Location.DriveAddress + 1
              "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
            }
            "Vault" {
              $vId = $_.VaultId
              $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
            "{0}: {1}" -f $_,$vName}
            default {"Lost in Space"}
          }
      }},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpPool += $subHead01 + "Expired Tapes in Media Pool: " + $mp.Name + $subHead02 + $expTapes
    }
  }
}

# Get Expired Tapes in each Vault
$bodyTpExpVlt = $null
If ($showExpTpVlt) {
  ForEach ($vlt in ($mediaVaults | Sort-Object Name)) {
    $expTapes = @($mediaTapes | Where-Object {($_.Location.VaultId -eq $vlt.Id -and $_.IsExpired -eq $True)})
    If ($expTapes.Count -gt 0) {
      $expTapes = $expTapes | Select-Object Name, Barcode,
      @{Name="Media Pool"; Expression = {
          $poolId = $_.MediaPoolId
          ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
      }},
      @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
      @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
      @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
      @{Name="Last Write"; Expression = {$_.LastWriteTime}} | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
      $bodyTpExpVlt += $subHead01 + "Expired Tapes in Vault: " + $vlt.Name + $subHead02 + $expTapes
    }
  }
}

# Get all Tapes written to within time frame
$bodyTpWrt = $null
If ($showTpWrt) {
  $expTapes = @($mediaTapes | Where-Object {$_.LastWriteTime -ge (Get-Date).AddHours(-$HourstoCheck)})
  If ($expTapes.Count -gt 0) {
    $expTapes = $expTapes | Select-Object Name, Barcode,
    @{Name="Media Pool"; Expression = {
        $poolId = $_.MediaPoolId
        ($mediaPools | Where-Object {$_.Id -eq $poolId}).Name
    }},
    @{Name="Media Set"; Expression = {$_.MediaSet}}, @{Name="Sequence #"; Expression = {$_.SequenceNumber}},
    @{Name="Location"; Expression = {
        switch ($_.Location) {
          "None" {"Offline"}
          "Slot" {
            $lId = $_.LibraryId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            [int]$slot = $_.SlotAddress + 1
            "{0} : {1} {2}" -f $lName,$_,$slot
          }
          "Drive" {
            $lId = $_.LibraryId
            $dId = $_.DriveId
            $lName = $($mediaLibs | Where-Object {$_.Id -eq $lId}).Name
            $dName = $($mediaDrives | Where-Object {$_.Id -eq $dId}).Name
            [int]$dNum = $_.Location.DriveAddress + 1
            "{0} : {1} {2} (Drive ID: {3})" -f $lName,$_,$dNum,$dName
          }
          "Vault" {
            $vId = $_.VaultId
            $vName = $($mediaVaults | Where-Object {$_.Id -eq $vId}).Name
          "{0}: {1}" -f $_,$vName}
          default {"Lost in Space"}
        }
    }},
    @{Name="Capacity (GB)"; Expression = {[Math]::Round([Decimal]$_.Capacity/1GB, 2)}},
    @{Name="Free (GB)"; Expression = {[Math]::Round([Decimal]$_.Free/1GB, 2)}},
    @{Name="Last Write"; Expression = {$_.LastWriteTime}},
    @{Name="Expiration Date"; Expression = {
        If ($(Get-Date $_.ExpirationDate) -lt $(Get-Date)) {
          "Expired"
        } Else {
          $_.ExpirationDate
        }
    }} | Sort-Object "Last Write" | ConvertTo-HTML -Fragment
    $bodyTpWrt = $subHead01 + "All Tapes Written" + $subHead02 + $expTapes
  }
}

# Get Agent Backup Summary Info
$bodySummaryEp = $null
If ($showSummaryEp) {
  $vbrEpHash = @{
    "Sessions" = If ($sessListEp) {@($sessListEp).Count} Else {0}
    "Successful" = @($successSessionsEp).Count
    "Warning" = @($warningSessionsEp).Count
    "Fails" = @($failsSessionsEp).Count
    "Running" = @($runningSessionsEp).Count
  }
  $vbrEPObj = New-Object -TypeName PSObject -Property $vbrEpHash
  If ($onlyLastEp) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummaryEp =  $vbrEPObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummaryEp = $arrSummaryEp | ConvertTo-HTML -Fragment
  If ($arrSummaryEp.Failures -gt 0) {
      $summaryEpHead = $subHead01err
  } ElseIf ($arrSummaryEp.Warnings -gt 0) {
      $summaryEpHead = $subHead01war
  } ElseIf ($arrSummaryEp.Successful -gt 0) {
      $summaryEpHead = $subHead01suc
  } Else {
      $summaryEpHead = $subHead01
  }
  $bodySummaryEp = $summaryEpHead + "Agent Backup Results Summary" + $subHead02 + $bodySummaryEp
}

# Get Agent Backup Job Status
$bodyJobsEp = $null
If ($showJobsEp) {
  If ($allJobsEp.count -gt 0) {
    $bodyJobsEp = $allJobsEp | Sort-Object Name | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Description"; Expression = {$_.Description}},
      @{Name="Enabled"; Expression = {$_.JobEnabled}},
      @{Name="Status"; Expression = {(Get-VBRComputerBackupJobSession -Name $_.Name)[0].state}},
      @{Name="Target Repo"; Expression = {$_.BackupRepository.Name}},
      @{Name="Next Run"; Expression = {
                If ($_.ScheduleEnabled -eq $false) {"<not scheduled>"}
                Else {(Get-VBRJobScheduleOptions -Job $_).nextrun}}},
      @{Name="Last Result"; Expression = {(Get-VBRComputerBackupJobSession -Name $_.Name)[0].result}} | ConvertTo-HTML -Fragment
    $bodyJobsEp = $subHead01 + "Agent Backup Job Status" + $subHead02 + $bodyJobsEp
  }
}

# Get Agent Backup Job Size
$bodyJobSizeEp = $null
If ($showBackupSizeEp) {
  If ($backupsEp.count -gt 0) {
    $bodyJobSizeEp = Get-BackupSize -backups $backupsEp | Sort-Object JobName | Select-Object @{Name="Job Name"; Expression = {$_.JobName}},
      @{Name="VM Count"; Expression = {$_.VMCount}},
      @{Name="Repository"; Expression = {$_.Repo}},
      @{Name="Data Size (GB)"; Expression = {$_.DataSize}},
      @{Name="Backup Size (GB)"; Expression = {$_.BackupSize}} | ConvertTo-HTML -Fragment
    $bodyJobSizeEp = $subHead01 + "Agent Backup Job Size" + $subHead02 + $bodyJobSizeEp
  }
}

# Get Agent Backup Sessions
$bodyAllSessEp = @()
$arrAllSessEp = @()
If ($showAllSessEp) {
  If ($sessListEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $arrAllSessEp += $sessListEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="State"; Expression = {$_.State}},@{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},
        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    }
    $bodyAllSessEp = $arrAllSessEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllSessEp.Result -match "Failed") {
        $allSessEpHead = $subHead01err
      } ElseIf ($arrAllSessEp.Result -match "Warning") {
        $allSessEpHead = $subHead01war
      } ElseIf ($arrAllSessEp.Result -match "Success") {
        $allSessEpHead = $subHead01suc
      } Else {
        $allSessEpHead = $subHead01
      }
    $bodyAllSessEp = $allSessEpHead + "Agent Backup Sessions" + $subHead02 + $bodyAllSessEp
  }
}

# Get Running Agent Backup Jobs
$bodyRunningEp = @()
If ($showRunningEp) {
  If ($runningSessionsEp.count -gt 0) {
    Foreach($job in $allJobsEp) {
      $bodyRunningEp += $runningSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}}
    }
    $bodyRunningEp = $bodyRunningEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyRunningEp = $subHead01 + "Running Agent Backup Jobs" + $subHead02 + $bodyRunningEp
  }
}

# Get Agent Backup Sessions with Warnings or Failures
$bodySessWFEp = @()
$arrSessWFEp = @()
If ($showWarnFailEp) {
  $sessWFEp = @($warningSessionsEp + $failsSessionsEp)
  If ($sessWFEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerWFEp = "Agent Backup Jobs with Warnings or Failures"
    } Else {
      $headerWFEp = "Agent Backup Sessions with Warnings or Failures"
    }
    Foreach($job in $allJobsEp) {
      $arrSessWFEp += $sessWFEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessWFEp = $arrSessWFEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrSessWFEp.Result -match "Failed") {
        $sessWFEpHead = $subHead01err
      } ElseIf ($arrSessWFEp.Result -match "Warning") {
        $sessWFEpHead = $subHead01war
      } ElseIf ($arrSessWFEp.Result -match "Success") {
        $sessWFEpHead = $subHead01suc
      } Else {
        $sessWFEpHead = $subHead01
      }
    $bodySessWFEp = $sessWFEpHead + $headerWFEp + $subHead02 + $bodySessWFEp
  }
}

# Get Successful Agent Backup Sessions
$bodySessSuccEp = @()
If ($showSuccessEp) {
  If ($successSessionsEp.count -gt 0) {
    If ($onlyLastEp) {
      $headerSuccEp = "Successful Agent Backup Jobs"
    } Else {
      $headerSuccEp = "Successful Agent Backup Sessions"
    }
    Foreach($job in $allJobsEp) {
      $bodySessSuccEp += $successSessionsEp | Where-Object {$_.JobId -eq $job.Id} | Select-Object @{Name="Job Name"; Expression = {$job.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}}, @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result
    }
    $bodySessSuccEp = $bodySessSuccEp | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodySessSuccEp = $subHead01suc + $headerSuccEp + $subHead02 + $bodySessSuccEp
  }
}

# Get SureBackup Summary Info
$bodySummarySb = $null
If ($showSummarySb) {
  $vbrMasterHash = @{
    "Sessions" = If ($sessListSb) {@($sessListSb).Count} Else {0}
    "Successful" = @($successSessionsSb).Count
    "Warning" = @($warningSessionsSb).Count
    "Fails" = @($failsSessionsSb).Count
    "Running" = @($runningSessionsSb).Count
  }
  $vbrMasterObj = New-Object -TypeName PSObject -Property $vbrMasterHash
  If ($onlyLastSb) {
    $total = "Jobs Run"
  } Else {
    $total = "Total Sessions"
  }
  $arrSummarySb =  $vbrMasterObj | Select-Object @{Name=$total; Expression = {$_.Sessions}},
    @{Name="Running"; Expression = {$_.Running}}, @{Name="Successful"; Expression = {$_.Successful}},
    @{Name="Warnings"; Expression = {$_.Warning}}, @{Name="Failures"; Expression = {$_.Fails}}
  $bodySummarySb = $arrSummarySb | ConvertTo-HTML -Fragment
  If ($arrSummarySb.Failures -gt 0) {
      $summarySbHead = $subHead01err
  } ElseIf ($arrSummarySb.Warnings -gt 0) {
      $summarySbHead = $subHead01war
  } ElseIf ($arrSummarySb.Successful -gt 0) {
      $summarySbHead = $subHead01suc
  } Else {
      $summarySbHead = $subHead01
  }
  $bodySummarySb = $summarySbHead + "SureBackup Results Summary" + $subHead02 + $bodySummarySb
}

# Get SureBackup Job Status
$bodyJobsSb = $null
If ($showJobsSb) {
  If ($allJobsSb.count -gt 0) {
    $bodyJobsSb = @()
    Foreach($SbJob in $allJobsSb) {
      $bodyJobsSb += $SbJob | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Enabled"; Expression = {$_.IsScheduleEnabled}},
        @{Name="Status"; Expression = {
          If ($_.GetLastState() -eq "Working") {
            $currentSess = $_.FindLastSession()
            $csessPercent = $currentSess.CompletionPercentage
            $cStatus = "$($csessPercent)% completed"
            $cStatus
          } Else {
            $_.GetLastState()
          }
        }},
        @{Name="Virtual Lab"; Expression = {$(Get-VSBVirtualLab | Where-Object {$_.Id -eq $SbJob.VirtualLabId}).Name}},
        @{Name="Linked Jobs"; Expression = {$($_.GetLinkedJobs()).Name -join ","}},
        @{Name="Next Run"; Expression = {
          If ($_.IsScheduleEnabled -eq $false) {"<Disabled>"}
          ElseIf ($_.JobOptions.RunManually) {"<not scheduled>"}
          ElseIf ($_.ScheduleOptions.IsContinuous) {"<Continuous>"}
          ElseIf ($_.ScheduleOptions.OptionsScheduleAfterJob.IsEnabled) {"After [" + $(($allJobs + $allJobsTp) | Where-Object {$_.Id -eq $SbJob.Info.ParentScheduleId}).Name + "]"}
          Else {$_.ScheduleOptions.NextRun}}},
        @{Name="Last Result"; Expression = {If ($_.GetLastResult() -eq "None"){""}Else{$_.GetLastResult()}}}
    }
    $bodyJobsSb = $bodyJobsSb | Sort-Object "Next Run" | ConvertTo-HTML -Fragment
    $bodyJobsSb = $subHead01 + "SureBackup Job Status" + $subHead02 + $bodyJobsSb
  }
}

# Get SureBackup Sessions
$bodyAllSessSb = $null
If ($showAllSessSb) {
  If ($sessListSb.count -gt 0) {
    $arrAllSessSb = $sessListSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="State"; Expression = {$_.State}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {If ($_.EndTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.EndTime}}},

        @{Name="Duration (HH:MM:SS)"; Expression = {
          If ($_.EndTime -eq "1/1/1900 12:00:00 AM") {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))
          } Else {
            Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)
          }
        }}, Result
    $bodyAllSessSb = $arrAllSessSb | ConvertTo-HTML -Fragment
    If ($arrAllSessSb.Result -match "Failed") {
        $allSessSbHead = $subHead01err
      } ElseIf ($arrAllSessSb.Result -match "Warning") {
        $allSessSbHead = $subHead01war
      } ElseIf ($arrAllSessSb.Result -match "Success") {
        $allSessSbHead = $subHead01suc
      } Else {
        $allSessSbHead = $subHead01
      }
    $bodyAllSessSb = $allSessSbHead + "SureBackup Sessions" + $subHead02 + $bodyAllSessSb
    }
}

# Get Running SureBackup Jobs
$bodyRunningSb = $null
If ($showRunningSb) {
  If ($runningSessionsSb.count -gt 0) {
    $bodyRunningSb = $runningSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
      @{Name="Start Time"; Expression = {$_.CreationTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $(Get-Date))}},
      @{Name="% Complete"; Expression = {$_.Progress}} | ConvertTo-HTML -Fragment
    $bodyRunningSb = $subHead01 + "Running SureBackup Jobs" + $subHead02 + $bodyRunningSb
  }
}

# Get SureBackup Sessions with Warnings or Failures
$bodySessWFSb = $null
If ($showWarnFailSb) {
  $sessWF = @($warningSessionsSb + $failsSessionsSb)
  If ($sessWF.count -gt 0) {
    If ($onlyLastSb) {
      $headerWF = "SureBackup Jobs with Warnings or Failures"
    } Else {
      $headerWF = "SureBackup Sessions with Warnings or Failures"
    }
    $arrSessWFSb = $sessWF | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}}, Result
    $bodySessWFSb = $arrSessWFSb | ConvertTo-HTML -Fragment
    If ($arrSessWFSb.Result -match "Failed") {
        $sessWFSbHead = $subHead01err
      } ElseIf ($arrSessWFSb.Result -match "Warning") {
        $sessWFSbHead = $subHead01war
      } ElseIf ($arrSessWFSb.Result -match "Success") {
        $sessWFSbHead = $subHead01suc
      } Else {
        $sessWFSbHead = $subHead01
      }
    $bodySessWFSb = $sessWFSbHead + $headerWF + $subHead02 + $bodySessWFSb
    }
}

# Get Successful SureBackup Sessions
$bodySessSuccSb = $null
If ($showSuccessSb) {
  If ($successSessionsSb.count -gt 0) {
    If ($onlyLastSb) {
      $headerSucc = "Successful SureBackup Jobs"
    } Else {
      $headerSucc = "Successful SureBackup Sessions"
    }
    $bodySessSuccSb = $successSessionsSb | Sort-Object Creationtime | Select-Object @{Name="Job Name"; Expression = {$_.Name}},
        @{Name="Start Time"; Expression = {$_.CreationTime}},
        @{Name="Stop Time"; Expression = {$_.EndTime}},
        @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.CreationTime $_.EndTime)}},
        Result | ConvertTo-HTML -Fragment
    $bodySessSuccSb = $subHead01suc + $headerSucc + $subHead02 + $bodySessSuccSb
  }
}

## Gathering tasks after session info has been recorded due to Veeam issue
# Gather all SureBackup Tasks from Sessions within time frame
$taskListSb = @()
$taskListSb += $sessListSb | Get-VSBTaskSession
$successTasksSb = @($taskListSb | Where-Object {$_.Info.Result -eq "Success"})
$wfTasksSb = @($taskListSb | Where-Object {$_.Info.Result -match "Warning|Failed"})
$runningTasksSb = @()
$runningTasksSb += $runningSessionsSb | Get-VSBTaskSession | Where-Object {$_.Status -ne "Stopped"}

# Get SureBackup Tasks
$bodyAllTasksSb = $null
If ($showAllTasksSb) {
  If ($taskListSb.count -gt 0) {
    $arrAllTasksSb = $taskListSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Status"; Expression = {$_.Status}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM"){"-"} Else {$_.Info.FinishTime}}},
      @{Name="Duration (HH:MM:SS)"; Expression = {
        If ($_.Info.FinishTime -eq "1/1/1900 12:00:00 AM") {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))
        } Else {
          Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)
        }
      }},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {
          If ($_.Info.Result -eq "notrunning") {
            "None"
          } Else {
            $_.Info.Result
          }
      }}
    $bodyAllTasksSb = $arrAllTasksSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrAllTasksSb.Result -match "Failed") {
        $allTasksSbHead = $subHead01err
      } ElseIf ($arrAllTasksSb.Result -match "Warning") {
        $allTasksSbHead = $subHead01war
      } ElseIf ($arrAllTasksSb.Result -match "Success") {
        $allTasksSbHead = $subHead01suc
      } Else {
        $allTasksSbHead = $subHead01
      }
    $bodyAllTasksSb = $allTasksSbHead + "SureBackup Tasks" + $subHead02 + $bodyAllTasksSb
  }
}

# Get Running SureBackup Tasks
$bodyTasksRunningSb = $null
If ($showRunningTasksSb) {
  If ($runningTasksSb.count -gt 0) {
    $bodyTasksRunningSb = $runningTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $(Get-Date))}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      Status | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTasksRunningSb = $subHead01 + "Running SureBackup Tasks" + $subHead02 + $bodyTasksRunningSb
  }
}

# Get SureBackup Tasks with Warnings or Failures
$bodyTaskWFSb = $null
If ($showTaskWFSb) {
  If ($wfTasksSb.count -gt 0) {
    $arrTaskWFSb = $wfTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}}
    $bodyTaskWFSb = $arrTaskWFSb | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    If ($arrTaskWFSb.Result -match "Failed") {
        $taskWFSbHead = $subHead01err
      } ElseIf ($arrTaskWFSb.Result -match "Warning") {
        $taskWFSbHead = $subHead01war
      } ElseIf ($arrTaskWFSb.Result -match "Success") {
        $taskWFSbHead = $subHead01suc
      } Else {
        $taskWFSbHead = $subHead01
      }
    $bodyTaskWFSb = $taskWFSbHead + "SureBackup Tasks with Warnings or Failures" + $subHead02 + $bodyTaskWFSb
  }
}

# Get Successful SureBackup Tasks
$bodyTaskSuccSb = $null
If ($showTaskSuccessSb) {
  If ($successTasksSb.count -gt 0) {
    $bodyTaskSuccSb = $successTasksSb | Select-Object @{Name="VM Name"; Expression = {$_.Name}},
      @{Name="Job Name"; Expression = {$_.JobSession.JobName}},
      @{Name="Start Time"; Expression = {$_.Info.StartTime}},
      @{Name="Stop Time"; Expression = {$_.Info.FinishTime}},
      @{Name="Duration (HH:MM:SS)"; Expression = {Get-Duration -ts $(New-TimeSpan $_.Info.StartTime $_.Info.FinishTime)}},
      @{Name="Heartbeat Test"; Expression = {$_.HeartbeatStatus}},
      @{Name="Ping Test"; Expression = {$_.PingStatus}},
      @{Name="Script Test"; Expression = {$_.TestScriptStatus}},
      @{Name="Validation Test"; Expression = {$_.VadiationTestStatus}},
      @{Name="Result"; Expression = {$_.Info.Result}} | Sort-Object "Start Time" | ConvertTo-HTML -Fragment
    $bodyTaskSuccSb = $subHead01suc + "Successful SureBackup Tasks" + $subHead02 + $bodyTaskSuccSb
  }
}

# Get Configuration Backup Summary Info
$bodySummaryConfig = $null
If ($showSummaryConfig) {
  $vbrConfigHash = @{
    "Enabled" = $configBackup.Enabled
    "Status" = $configBackup.LastState
    "Target" = $configBackup.Target
    "Schedule" = $configBackup.ScheduleOptions
    "Restore Points" = $configBackup.RestorePointsToKeep
    "Encrypted" = $configBackup.EncryptionOptions.Enabled
    "Last Result" = $configBackup.LastResult
    "Next Run" = $configBackup.NextRun
  }
  $vbrConfigObj = New-Object -TypeName PSObject -Property $vbrConfigHash
  $bodySummaryConfig = $vbrConfigObj | Select-Object Enabled, Status, Target, Schedule, "Restore Points", "Next Run", Encrypted, "Last Result" | ConvertTo-HTML -Fragment
  If ($configBackup.LastResult -eq "Warning" -or !$configBackup.Enabled) {
    $configHead = $subHead01war
  } ElseIf ($configBackup.LastResult -eq "Success") {
    $configHead = $subHead01suc
  } ElseIf ($configBackup.LastResult -eq "Failed") {
    $configHead = $subHead01err
  } Else {
    $configHead = $subHead01
  }
  $bodySummaryConfig = $configHead + "Configuration Backup Status" + $subHead02 + $bodySummaryConfig
}

# Get Proxy Info
$bodyProxy = $null
If ($showProxy) {
  If ($proxyList.count -gt 0) {
    $arrProxy = $proxyList | Get-VBRProxyInfo | Select-Object @{Name="Proxy Name"; Expression = {$_.ProxyName}},
      @{Name="Transport Mode"; Expression = {$_.tMode}}, @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Proxy Host"; Expression = {$_.RealName}}, @{Name="Host Type"; Expression = {$_.pType}},
      Enabled, @{Name="IP Address"; Expression = {$_.IP}},
      @{Name="RT (ms)"; Expression = {$_.Response}}, Status
    $bodyProxy = $arrProxy | Sort-Object "Proxy Host" |  ConvertTo-HTML -Fragment
    If ($arrProxy.Status -match "Dead") {
      $proxyHead = $subHead01err
    } ElseIf ($arrProxy -match "Alive") {
      $proxyHead = $subHead01suc
    } Else {
      $proxyHead = $subHead01
    }
    $bodyProxy = $proxyHead + "Proxy Details" + $subHead02 + $bodyProxy
  }
}

# Get Repository Info
$bodyRepo = $null
If ($showRepo) {
  If ($repoList.count -gt 0) {
    $arrRepo = $repoList | Get-VBRRepoInfo | Select-Object @{Name="Repository Name"; Expression = {$_.Target}},
      @{Name="Type"; Expression = {$_.rType}},
      @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
      @{Name="Host"; Expression = {$_.RepoHost}},
      @{Name="Path"; Expression = {$_.Storepath}},
      @{Name="Backups (GB)"; Expression = {$_.StorageBackup}},
      @{Name="Other data (GB)"; Expression = {$_.StorageOther}},
      @{Name="Free (GB)"; Expression = {$_.StorageFree}},
      @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0 -and $_.rtype -ne "SAN Snapshot")  {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}
      }
    $bodyRepo = $arrRepo | Sort-Object "Repository Name" | ConvertTo-HTML -Fragment
    If ($arrRepo.status -match "Critical") {
      $repoHead = $subHead01err
    } ElseIf ($arrRepo.status -match "Warning|Unknown") {
      $repoHead = $subHead01war
    } ElseIf ($arrRepo.status -match "OK") {
      $repoHead = $subHead01suc
    } Else {
      $repoHead = $subHead01
    }
    $bodyRepo = $repoHead + "Repository Details" + $subHead02 + $bodyRepo
  }
}
# Get Scale Out Repository Info
$bodySORepo = $null
If ($showRepo) {
  If ($repoListSo.count -gt 0) {
    $arrSORepo = $repoListSo | Get-VBRSORepoInfo | Select-Object @{Name="Scale Out Repository Name"; Expression = {$_.SOTarget}},
      @{Name="Member Name"; Expression = {$_.Target}},
	  @{Name="Type"; Expression = {$_.rType}},
      @{Name="Max Tasks"; Expression = {$_.MaxTasks}},
	  @{Name="Host"; Expression = {$_.RepoHost}},
      @{Name="Path"; Expression = {$_.Storepath}},
      @{Name="Backups (GB)"; Expression = {$_.StorageBackup}},
      @{Name="Other data (GB)"; Expression = {$_.StorageOther}},
	  @{Name="Free (GB)"; Expression = {$_.StorageFree}},
      @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
	  @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $repoCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
        ElseIf ($_.FreePercentage -lt $repoWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}}

      }
    $bodySORepo = $arrSORepo | Sort-Object "Scale Out Repository Name", "Member Repository Name" | ConvertTo-HTML -Fragment
    If ($arrSORepo.status -match "Critical") {
      $sorepoHead = $subHead01err
    } ElseIf ($arrSORepo.status -match "Warning|Unknown") {
      $sorepoHead = $subHead01war
    } ElseIf ($arrSORepo.status -match "OK") {
      $sorepoHead = $subHead01suc
    } Else {
      $sorepoHead = $subHead01
    }
    $bodySORepo = $sorepoHead + "Scale Out Repository Details" + $subHead02 + $bodySORepo
  }
}

# Get Repository Agent User Permissions
$bodyRepoPerms = $null
If ($showRepoPerms){
  If ($repoList.count -gt 0 -or $repoListSo.count -gt 0) {
    $bodyRepoPerms = Get-RepoPermission | Select-Object Name, "Encryption Enabled", "Permission Type", Users | Sort-Object Name | ConvertTo-HTML -Fragment
    $bodyRepoPerms = $subHead01 + "Repository Permissions for Agent Jobs" + $subHead02 + $bodyRepoPerms
  }
}

# Get Replica Target Info
$bodyReplica = $null
If ($showReplicaTarget) {
  If ($allJobsRp.count -gt 0) {
    $repTargets = $allJobsRp | Get-VBRReplicaTarget | Select-Object @{Name="Replica Target"; Expression = {$_.Target}}, Datastore,
      @{Name="Free (GB)"; Expression = {$_.StorageFree}}, @{Name="Total (GB)"; Expression = {$_.StorageTotal}},
      @{Name="Free (%)"; Expression = {$_.FreePercentage}},
      @{Name="Status"; Expression = {
        If ($_.FreePercentage -lt $replicaCritical) {"Critical"}
        ElseIf ($_.StorageTotal -eq 0)  {"Warning"}
        ElseIf ($_.FreePercentage -lt $replicaWarn) {"Warning"}
        ElseIf ($_.FreePercentage -eq "Unknown") {"Unknown"}
        Else {"OK"}
        }
      } | Sort-Object "Replica Target"
    $bodyReplica = $repTargets | ConvertTo-HTML -Fragment
    If ($repTargets.status -match "Critical") {
      $reptarHead = $subHead01err
    } ElseIf ($repTargets.status -match "Warning|Unknown") {
      $reptarHead = $subHead01war
    } ElseIf ($repTargets.status -match "OK") {
      $reptarHead = $subHead01suc
    } Else {
      $reptarHead = $subHead01
    }
    $bodyReplica = $reptarHead + "Replica Target Details" + $subHead02 + $bodyReplica
  }
}

# Get Veeam Services Info
$bodyServices = $null
If ($showServices) {
  $vServers = Get-VeeamWinServer
  $vServices = Get-VeeamService $vServers
  If ($hideRunningSvc) {$vServices = $vServices | Where-Object {$_.Status -ne "Running"}}
  If ($vServices.count -gt 0) {
    $vServices = $vServices | Select-Object "Server Name", "Service Name",
      @{Name="Status"; Expression = {If ($_.Status -eq "Stopped"){"Not Running"} Else {$_.Status}}}
    $bodyServices = $vServices | Sort-Object "Server Name", "Service Name" | ConvertTo-HTML -Fragment
    If ($vServices.status -match "Not Running") {
      $svcHead = $subHead01err
    } ElseIf ($vServices.status -notmatch "Running") {
      $svcHead = $subHead01war
    } ElseIf ($vServices.status -match "Running") {
      $svcHead = $subHead01suc
    } Else {
      $svcHead = $subHead01
    }
    $bodyServices = $svcHead + "Veeam Services (Windows)" + $subHead02 + $bodyServices
  }
}

#region license info
# Get License Info
$bodyLicense = $null
If ($showLicExp) {
  $arrLicense = Get-VeeamSupportDate $vbrServer | Select-Object @{Name = "Type"; Expression = { $_.LicType } },
    @{Name="Expiry Date"; Expression = {$_.ExpDate}},
    @{Name="Days Remaining"; Expression = {$_.DaysRemain}}, `
    @{Name="Status"; Expression = {
      If ($_.LicType -eq "Evaluation") {"OK"}
      ElseIf ($_.DaysRemain -lt $licenseCritical) {"Critical"}
      ElseIf ($_.DaysRemain -lt $licenseWarn) {"Warning"}
      ElseIf ($_.DaysRemain -eq "Failed") {"Failed"}
      Else {"OK"}}
    }
  $bodyLicense = $arrLicense | ConvertTo-HTML -Fragment
    If ($arrLicense.Type -eq "Evaluation") {
        $licHead = $subHead01inf
    } Else {
      If ($arrLicense.Status -eq "OK") {
        $licHead = $subHead01suc
      } ElseIf ($arrLicense.Status -eq "Warning") {
        $licHead = $subHead01war
      } Else {
        $licHead = $subHead01err
      }
    }
  $bodyLicense = $licHead + "License/Support Renewal Date" + $subHead02 + $bodyLicense
}
#endregion

#region Combine HTML Output
$htmlOutput = $headerObj + $bodyTop + $bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb

If ($bodySummaryProtect + $bodySummaryBK + $bodySummaryRp + $bodySummaryBc + $bodySummaryTp + $bodySummaryEp + $bodySummarySb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMissing + $bodyWarning + $bodySuccess

If ($bodyMissing + $bodySuccess + $bodyWarning) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyMultiJobs

If ($bodyMultiJobs) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBk + $bodyJobSizeBk + $bodyFileJobsBk + $bodyFileJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk

If ($bodyJobsBk + $bodyJobSizeBk + $bodyAllSessBk + $bodyAllTasksBk + $bodyRunningBk + $bodyTasksRunningBk + $bodySessWFBk + $bodyTaskWFBk + $bodySessSuccBk + $bodyTaskSuccBk) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyRestoRunVM + $bodyRestoreVM

If ($bodyRestoRunVM + $bodyRestoreVM) {
  $htmlOutput += $HTMLbreak
  }

$htmlOutput += $bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp

If ($bodyJobsRp + $bodyAllSessRp + $bodyAllTasksRp + $bodyRunningRp + $bodyTasksRunningRp + $bodySessWFRp + $bodyTaskWFRp + $bodySessSuccRp + $bodyTaskSuccRp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc

If ($bodyJobsBc + $bodyJobSizeBc + $bodyAllSessBc + $bodyAllTasksBc + $bodySessIdleBc + $bodyTasksPendingBc + $bodyRunningBc + $bodyTasksRunningBc + $bodySessWFBc + $bodyTaskWFBc + $bodySessSuccBc + $bodyTaskSuccBc) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp

If ($bodyJobsTp + $bodyAllSessTp + $bodyAllTasksTp + $bodyWaitingTp + $bodySessIdleTp + $bodyTasksPendingTp + $bodyRunningTp + $bodyTasksRunningTp + $bodySessWFTp + $bodyTaskWFTp + $bodySessSuccTp + $bodyTaskSuccTp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt

If ($bodyTapes + $bodyTpPool + $bodyTpVlt + $bodyExpTp + $bodyTpExpPool + $bodyTpExpVlt + $bodyTpWrt) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp

If ($bodyJobsEp + $bodyJobSizeEp + $bodyAllSessEp + $bodyRunningEp + $bodySessWFEp + $bodySessSuccEp) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb

If ($bodyJobsSb + $bodyAllSessSb + $bodyAllTasksSb + $bodyRunningSb + $bodyTasksRunningSb + $bodySessWFSb + $bodyTaskWFSb + $bodySessSuccSb + $bodyTaskSuccSb) {
  $htmlOutput += $HTMLbreak
}

$htmlOutput += $bodySummaryConfig + $bodyProxy + $bodyRepo + $bodySORepo + $bodyRepoPerms + $bodyReplica + $bodyServices + $bodyLicense + $footerObj

# Fix Details
$htmlOutput = $htmlOutput.Replace("ZZbrZZ","<br />")
# Remove trailing HTMLbreak
$htmlOutput = $htmlOutput.Replace("$($HTMLbreak + $footerObj)","$($footerObj)")
# Add color to output depending on results
#Green
$htmlOutput = $htmlOutput.Replace("<td>Running<","<td style=""color: #00b051;"">Running<")
$htmlOutput = $htmlOutput.Replace("<td>OK<","<td style=""color: #00b051;"">OK<")
$htmlOutput = $htmlOutput.Replace("<td>Alive<","<td style=""color: #00b051;"">Alive<")
$htmlOutput = $htmlOutput.Replace("<td>Success<","<td style=""color: #00b051;"">Success<")
#Yellow
$htmlOutput = $htmlOutput.Replace("<td>Warning<","<td style=""color: #ffc000;"">Warning<")
#Red
$htmlOutput = $htmlOutput.Replace("<td>Not Running<","<td style=""color: #ff0000;"">Not Running<")
$htmlOutput = $htmlOutput.Replace("<td>Failed<","<td style=""color: #ff0000;"">Failed<")
$htmlOutput = $htmlOutput.Replace("<td>Critical<","<td style=""color: #ff0000;"">Critical<")
$htmlOutput = $htmlOutput.Replace("<td>Dead<","<td style=""color: #ff0000;"">Dead<")
# Color Report Header and Tag Email Subject
If ($htmlOutput -match "#FB9895") {
  # If any errors paint report header red
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#FB9895")
  $emailSubject = "[Failed] $emailSubject"
} ElseIf ($htmlOutput -match "#ffd96c") {
  # If any warnings paint report header yellow
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#ffd96c")
  $emailSubject = "[Warning] $emailSubject"
} ElseIf ($htmlOutput -match "#00b050") {
  # If any success paint report header green
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#00b050")
  $emailSubject = "[Success] $emailSubject"
} Else {
  # Else paint gray
  $htmlOutput = $htmlOutput.Replace("ZZhdbgZZ","#626365")
}
#endregion

#region Output
# Send Report via Email
If ($sendEmail) {
  $smtp = New-Object System.Net.Mail.SmtpClient($emailHost, $emailPort)
  $smtp.Credentials = New-Object System.Net.NetworkCredential($emailUser, $emailPass)
  $smtp.EnableSsl = $emailEnableSSL
  $msg = New-Object System.Net.Mail.MailMessage($emailFrom, $emailTo)
  $msg.Subject = $emailSubject
  If ($emailAttach) {
    $body = "Veeam Report Attached"
    $msg.Body = $body
    $tempFile = "$env:TEMP\$($rptTitle)_$(Get-Date -format yyyyMMdd_HHmmss).htm"
    $htmlOutput | Out-File $tempFile
    $attachment = New-Object System.Net.Mail.Attachment $tempFile
    $msg.Attachments.Add($attachment)
  } Else {
    $body = $htmlOutput
    $msg.Body = $body
    $msg.isBodyhtml = $true
  }
  $smtp.send($msg)
  If ($emailAttach) {
    $attachment.dispose()
    Remove-Item $tempFile
  }
}

# Save HTML Report to File
If ($saveHTML) {
  $htmlOutput | Out-File $pathHTML
  If ($launchHTML) {
    Invoke-Item $pathHTML
  }
}
#endregion