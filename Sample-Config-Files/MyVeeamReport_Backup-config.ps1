#requires -Version 5.0
<#

    .SYNOPSIS
    Configuration file for My Veeam Report.

    .DESCRIPTION
    Put here all your report customization settings.

    .NOTES
    Authors: Bernhard Roth & Marco Horstmann
    Last Updated: 19 June 2023
    Version: 12.0.0.1
#>
<#
	.UPDATES
	Eamonn Deering
	Last Updated: 10 Sep 2023
	Config Version: 12.0.0.3
	The following two settings requires MyVeeamReport Version: 12.0.0.5 or above.
	
	SureBackup update
		Added: $ReportHasDataEmail
		Added: $ReportSBFailingOnly
	
	.VM Exclusion 
	Added: Text file import with VM exclusions
	
	.SUGGESTIONS
	I find setting $setCSVDelimiter to  "," works well for Excel
	
	#> 

<# 	
	.$ReportHasDataEmail
	ED Works with Veeam 11 and 12
	Last Updated:	11 Aug 2023 
	
	DESCRIPTION
	Only email a report if the report has data other than a blank html page.
	Experimental. 
	Tested with: 
	 "Get SureBackup Tasks with Warnings or Failures" using $showTaskWFSb and $ReportSBFailingOnly
	 "Get VMs Missing Backups" using $missingVMsEmail="SendEmailReport"

	For SureBackup or "VMs Missing" if the list has no failed VM's then don't send the email. 
	 
#>
#The default is "$ReportHasDataEmail = $false" to ignore this setting.
#$ReportHasDataEmail = $true
$ReportHasDataEmail = $false

<#
	.$ReportSBFailingOnly
	ED Works with Veeam 11 and 12
	Last Updated:	11 Aug 2023 
	
	DESCRIPTION
	Only report a SB failed VM if that same VM never passed in the time frame.
	If your only interested in SB VM's that fail consistently over the time frame. 
	Example. Your SB jobs test VM's every week. Your time span is one month. If a VM fails every week then send email. If you have a ticketing system then this could log a ticket for you. 
	If a VM passes even once durning the time span then don't send email. 
	Use with all other setting set to $false execpt "$showTaskWFSb = $true" and "$ReportSBFailingOnly = $true".

#>
#The default is "$ReportSBFailingOnly = $false" to ignore this setting.
#$ReportSBFailingOnly = $true
$ReportSBFailingOnly = $false


# VBR Server (Server Name, FQDN, IP or localhost)
$vbrServer = $env:computername
#$vbrServer = "lab-vbr01"
# Report mode (RPO) - valid modes: any number of hours, Weekly or Monthly
# 24, 48, "Weekly", "Monthly"
$reportMode = "Weekly"
# Report Title
$rptTitle = "My Veeam Report"
# Show VBR Server name in report header
$showVBR = $true
# HTML Report Width (Percent)
$rptWidth = 97
# HTML Table Odd Row color
$oddColor = "#f0f0f0"

# Location of Veeam Core dll  
$VeeamCorePath = "C:\Program Files\Veeam\Backup and Replication\Backup\Veeam.Backup.Core.dll"
#If you are connect remotely to VBR server you need to use another console.
#$VeeamCorePath = "C:\Program Files\Veeam\Backup and Replication\Console\Veeam.Backup.Core.dll"


# ED Moving reports into reports folder
if(!(Test-Path ".\Reports")){
	MD ".\Reports"}
# Save HTML output to a file
$saveHTML = $true
# HTML File output path and filename
$pathHTML = ".\Reports\MyVeeamReport_$(Get-Date -format yyyyMMdd_HHmmss).htm"
# Launch HTML file after creation
$launchHTML = $true

# Save CSV output to files
$saveCSV = $true
# CSV File output path and filename
$baseFilenameCSV = ".\Reports\MyVeeamReport_$(Get-Date -format yyyyMMdd_HHmmss)"
# Export All Tasks to CSV file
$exportAllTasksBkToCSV = $true
#Delimiter for CSV files
$setCSVDelimiter = ","


# Email configuration
$sendEmail = $false
$emailHost = "smtp.yourserver.com"
$emailPort = 25
$emailEnableSSL = $false
$emailUser = ""
$emailPass = ""
$emailFrom = "MyVeeamReport@yourdomain.com"
$emailTo = "you@youremail.com"
# Send HTML report as attachment (else HTML report is body)
$emailAttach = $false
# Email Subject 
$emailSubject = $rptTitle
# Append Report Mode to Email Subject E.g. My Veeam Report (Last 24 Hours)
$modeSubject = $true
# Append VBR Server name to Email Subject
$vbrSubject = $true
# Append Date and Time to Email Subject
$dtSubject = $false

#--------------------- Disable reports you do not need by setting them to "$false" below:                                                                                        
# Show VM Backup Protection Summary (across entire infrastructure)
$showSummaryProtect = $false
# Show VMs with No Successful Backups within RPO ($reportMode)
$showUnprotectedVMs = $false
# Show unprotected VMs for informational purposes only
$showUnprotectedVMsInfo = $false
# Show VMs with Successful Backups within RPO ($reportMode)
# Also shows VMs with Only Backups with Warnings within RPO ($reportMode)
$showProtectedVMs = $false
# Exclude VMs from Missing and Successful Backups sections
# $excludevms = @("vm1","vm2","*_replica")
$excludeVMs = @("")

##ED added text file import with exclusions.  
$ExcludeVMsSiteA = "C:\Scripts\Veeam\SiteA\Veeam-Exclude-VMs.txt"
$ExcludeVMsSiteB = "C:\Scripts\Veeam\SiteB\Veeam-Exclude-VMs.txt"

if(Test-Path $ExcludeVMsSiteA){
$ExcludeVMs += Get-Content $ExcludeVMsSiteA}
if(Test-Path $ExcludeVMsSiteB){
$ExcludeVMs += Get-Content $ExcludeVMsSiteB}
#$excludeVMs += @("*_replica*","*vLAB","*vLAN","*Template*")
$excludeVMs = $excludeVMs| Sort-Object -Property @{Expression={$_.Trim()}} -Unique
Write-Host " Excluded VM's "								   
$excludeVMs

# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) folder(s)
# $excludeFolder = @("folder1","folder2","*_testonly")
#$excludeFolder = @("")
$excludeFolder = @("DR-Testing","*Testing*","*Template*","*Powered-Off*","vm","vCLS","*To-Be-Deleted*","*vLab*")
# Exclude VMs from Missing and Successful Backups sections in the following (vCenter) datacenter(s)
# $excludeDC = @("dc1","dc2","dc*")
$excludeDC = @("")
# Exclude Templates from Missing and Successful Backups sections
$excludeTemp = $false

# Show VMs Backed Up by Multiple Jobs within time frame ($reportMode)
$showMultiJobs = $false

#<#
# Show Backup Session Summary
$showSummaryBk = $true
# Show Backup Job Status
$showJobsBk = $true
# Show File Backup Job Status
$showFileJobsBk = $true
# Show Backup Job Size (total)
$showBackupSizeBk = $true
# Show File Backup Job Size (total)
$showFileBackupSizeBk = $true
# Show detailed information for Backup Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBk = $false
# Show all Backup Sessions within time frame ($reportMode)
$showAllSessBk = $true
# Show all Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksBk = $true
# Show Running Backup Jobs
$showRunningBk = $true
# Show Running Backup Tasks
$showRunningTasksBk = $true
# Show Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBk = $true
# Show Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBk = $true
# Show Successful Backup Sessions within time frame ($reportMode)
$showSuccessBk = $true
# Show Successful Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBk = $true
# Only show last Session for each Backup Job
$onlyLastBk = $false
# Only report on the following Backup Job(s)
#$backupJob = @("Backup Job 1","Backup Job 3","Backup Job *")
$backupJob = @("")
#>

<#
# Show Running Restore VM Sessions
$showRestoRunVM = $true
# Show Completed Restore VM Sessions within time frame ($reportMode)
$showRestoreVM = $true

# Show Replication Session Summary
$showSummaryRp = $false
# Show Replication Job Status
$showJobsRp = $false
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $true
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $false
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $false
# Show Running Replication Jobs
$showRunningRp = $false
# Show Running Replication Tasks
$showRunningTasksRp = $false
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $false
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $false
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $false
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $false
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Backup Copy Session Summary
$showSummaryBc = $true
# Show Backup Copy Job Status
$showJobsBc = $true
# Show Backup Copy Job Size (total)
$showBackupSizeBc = $true
# Show detailed information for Backup Copy Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedBc = $true
# Show all Backup Copy Sessions within time frame ($reportMode)
$showAllSessBc = $true
# Show all Backup Copy Tasks from Sessions within time frame ($reportMode)
$showAllTasksBc = $true
# Show Idle Backup Copy Sessions
$showIdleBc = $true
# Show Pending Backup Copy Tasks
$showPendingTasksBc = $true
# Show Working Backup Copy Jobs
$showRunningBc = $true
# Show Working Backup Copy Tasks
$showRunningTasksBc = $true
# Show Backup Copy Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailBc = $true
# Show Backup Copy Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFBc = $true
# Show Successful Backup Copy Sessions within time frame ($reportMode)
$showSuccessBc = $true
# Show Successful Backup Copy Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessBc = $true
# Only show last Session for each Backup Copy Job
$onlyLastBc = $false
# Only report on the following Backup Copy Job(s)
#$bcopyJob = @("Backup Copy Job 1","Backup Copy Job 3","Backup Copy Job *")
$bcopyJob = @("")
#>

<#
# Show Tape Backup Session Summary
$showSummaryTp = $true
# Show Tape Backup Job Status
$showJobsTp = $true
# Show detailed information for Tape Backup Sessions (Avg Speed, Total(GB), Read(GB), Transferred(GB))
$showDetailedTp = $true
# Show all Tape Backup Sessions within time frame ($reportMode)
$showAllSessTp = $true
# Show all Tape Backup Tasks from Sessions within time frame ($reportMode)
$showAllTasksTp = $true
# Show Waiting Tape Backup Sessions
$showWaitingTp = $true
# Show Idle Tape Backup Sessions
$showIdleTp = $true
# Show Pending Tape Backup Tasks
$showPendingTasksTp = $true
# Show Working Tape Backup Jobs
$showRunningTp = $true
# Show Working Tape Backup Tasks
$showRunningTasksTp = $true
# Show Tape Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailTp = $true
# Show Tape Backup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFTp = $true
# Show Successful Tape Backup Sessions within time frame ($reportMode)
$showSuccessTp = $true
# Show Successful Tape Backup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessTp = $true
# Only show last Session for each Tape Backup Job
$onlyLastTp = $false
# Only report on the following Tape Backup Job(s)
#$tapeJob = @("Tape Backup Job 1","Tape Backup Job 3","Tape Backup Job *")
$tapeJob = @("")

# Show all Tapes
$showTapes = $true
# Show all Tapes by (Custom) Media Pool
$showTpMp = $true
# Show all Tapes by Vault
$showTpVlt = $true
# Show all Expired Tapes
$showExpTp = $true
# Show Expired Tapes by (Custom) Media Pool
$showExpTpMp = $true
# Show Expired Tapes by Vault
$showExpTpVlt = $true
# Show Tapes written to within time frame ($reportMode)
$showTpWrt = $true
#>

<#
# Show Agent Backup Session Summary
$showSummaryEp = $true
# Show Agent Backup Job Status
$showJobsEp = $true
# Show Agent Backup Job Size (total)
$showBackupSizeEp = $true
# Show all Agent Backup Sessions within time frame ($reportMode)
$showAllSessEp = $true
# Show Running Agent Backup jobs
$showRunningEp = $true
# Show Agent Backup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailEp = $true
# Show Successful Agent Backup Sessions within time frame ($reportMode)
$showSuccessEp = $true
# Only show last session for each Agent Backup Job
$onlyLastEp = $false
# Only report on the following Agent Backup Job(s)
#$epbJob = @("Agent Backup Job 1","Agent Backup Job 3","Agent Backup Job *")
$epbJob = @("")

# Show Configuration Backup Summary
$showSummaryConfig = $true
# Show Proxy Info
$showProxy = $true
# Show Repository Info
$showRepo = $true
# Show Repository Permissions for Agent Jobs
$showRepoPerms = $true
# Show Replica Target Info
$showReplicaTarget = $true
# Show Veeam Services Info (Windows Services)
$showServices = $true
# Show only Services that are NOT running
$hideRunningSvc = $true
# Show License expiry info
$showLicExp = $true


# Show SureBackup Session Summary
$showSummarySb = $true
# Show SureBackup Job Status
$showJobsSb = $true
# Show all SureBackup Sessions within time frame ($reportMode)
$showAllSessSb = $true
# Show all SureBackup Tasks from Sessions within time frame ($reportMode)
$showAllTasksSb = $true
# Show Running SureBackup Jobs
$showRunningSb = $true
# Show Running SureBackup Tasks
$showRunningTasksSb = $true
# Show SureBackup Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailSb = $true
# Show SureBackup Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFSb = $true
# Show Successful SureBackup Sessions within time frame ($reportMode)
$showSuccessSb = $true
# Show Successful SureBackup Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessSb = $true
# Only show last Session for each SureBackup Job
$onlyLastSb = $false
# Only report on the following SureBackup Job(s)
#$surebJob = @("SureBackup Job 1","SureBackup Job 3","SureBackup Job *")
$surebJob = @("")
#>
#Start of unchanged reports since version 9.5.3
# ED Note. Most of the following look to work in V11 and VBR v12.0.0.1420 as is.																																													 
<#
# Show Replication Session Summary
$showSummaryRp = $false
# Show Replication Job Status
$showJobsRp = $false
# Show detailed information for Replication Jobs/Sessions (Avg Speed, Total(GB), Processed(GB), Read(GB), Transferred(GB), Dedupe, Compression)
$showDetailedRp = $false
# Show all Replication Sessions within time frame ($reportMode)
$showAllSessRp = $false
# Show all Replication Tasks from Sessions within time frame ($reportMode)
$showAllTasksRp = $false
# Show Running Replication Jobs
$showRunningRp = $false
# Show Running Replication Tasks
$showRunningTasksRp = $false
# Show Replication Sessions w/Warnings or Failures within time frame ($reportMode)
$showWarnFailRp = $false
# Show Replication Tasks w/Warnings or Failures from Sessions within time frame ($reportMode)
$showTaskWFRp = $false
# Show Successful Replication Sessions within time frame ($reportMode)
$showSuccessRp = $false
# Show Successful Replication Tasks from Sessions within time frame ($reportMode)
$showTaskSuccessRp = $false
# Only show last session for each Replication Job
$onlyLastRp = $false
# Only report on the following Replication Job(s)
#$replicaJob = @("Replica Job 1","Replica Job 3","Replica Job *")
$replicaJob = @("")

# Show Running Restore VM Sessions
$showRestoRunVM = $false
# Show Completed Restore VM Sessions within time frame ($reportMode)
$showRestoreVM = $false

end of excluded unchanged reports since version 9.5.3 #>


# Highlighting Thresholds
# Repository Free Space Remaining %
$repoCritical = 10
$repoWarn = 20
# Replica Target Free Space Remaining %
$replicaCritical = 10
$replicaWarn = 20
# License Days Remaining
$licenseCritical = 30
$licenseWarn = 90
