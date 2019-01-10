<#
	References for the Windows Update API:
		https://docs.microsoft.com/en-us/windows/desktop/api/wuapi/
		https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa386400(v%3dvs.85)

	Parameters: (Can be given in any length, e.g. -ThresholdUpdate. Thanks PS!)
		ThresholdUpdateDays:		How many days should we look back for updates.
						If no updates are found within the threshold we consider the system out-of-date

		ThresholdFailedUpdateHours:	How many hours should we look back for updates.
						If you run this script daily you can set it to 23-24 hours for example.

		NoExit:				Controls wherever Powershell should hard exit with the given error code
						Just append if you don't want it to, no value needed

		NoCheckFailed:			Skips the check for failed updates. 
						Excludes error code 1002.

	Example usage:
		powershell -ExecutionPolicy ByPass -File CheckInstalledWindowsUpdates.ps1 -ThresholdUpdate 30 -ThresholdFailedUpdateHours 168 -NoExit

	Exit codes:
		0:				All systems nominal.
						Either there are successful updates or the installation / last feature update 
						(e.g. 1803->1809) was in the given time period.

		1001:				No updates have been installed in the given time period

		1002:				There have been failed updates in the last 24 hours. Supersedes 1001.

		1003:				Not a single successful update has been found in the logs. 
						Usually a sign for a borked Windows Update.

		1004:				The Windows Update log is completely empty.
						Another sign for a borked Windows Update.

		1005:				No connection could be made to the Windows Update service.
						Yet another sign for a borked Windows Update.
	
	Below are all default values aswell as all defined strings including their placeholders.
	Change them to fit your use case and/or language
#>

param(
	[switch]$NoExit				= $false,
	[switch]$NoCheckFailed			= $false,
	[int]$ThresholdUpdateDays		= 45,
	[int]$ThresholdFailedUpdateHours	= 23	
)
[string]$StringCannotConnectToWindowsUpdate	= "Cannot connect to Windows Update service"
[string]$StringFeatureUpdate			= "Last feature update was {0} days ago."
[string]$StringNoUpdatesFound			= "No updates found in update history."
[string]$StringNoSuccessfulUpdate		= "No successful update found in update history."
[string]$StringUpdatesWaitingForInstallation	= "{0} updates waiting for installation/reboot:`n{1}"
[string]$StringErrorWhileInstallingUpdates	= "Error while installing {0} updates:`n{1}"
[string]$StringLastUpdateInstalledon		= "Last update installed on {0} ({1})"
[string]$StringNoUpdatesInstalledInTimePeriod	= "No updates have been installed within the last $ThresholdUpdateDays days"
[string]$StringNoUpdatesFailed			= "No updates failed to install."
[string]$StringUpdatesInstalledinTimePeriod	= "{0} installed updates in the last $ThresholdUpdateDays days:`n{1}"

<# -- -- -- -- -- -- -- -- #>

function ScriptExit{
	param([int]$ScriptExitCode)
	if (!$NoExit){
		$host.SetShouldExit($ScriptExitCode)
	}
	exit $ScriptExitCode
}

function Notify{
	param(
		[switch]$Error,
		[string]$Message
	)
	if($Error){
		Write-Host $Message -ForegroundColor red -BackgroundColor black
	}else{
		Write-Host $Message -ForegroundColor yellow
	}
}

try{
	$session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$ComputerName))
	$updateSearcher = $session.CreateUpdateSearcher()
	$updateHistoryCount = $updateSearcher.GetTotalHistoryCount()
}catch [Exception]{
		Notify -Error $StringCannotConnectToWindowsUpdate
		Notify -Error $_.Exception.Message
		ScriptExit 1005
}

if ( $updateHistoryCount -le 0 ){

	[Nullable[datetime]]$FeatureUpdateInstallDate 
	Try{
		[datetime]$FeatureUpdateInstallDate = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).InstallDate
	}Catch [Exception]{
		[datetime]$FeatureUpdateInstallDate = ((get-date -year 1970 -month 1 -day 1 -hour 0 -minute 0 -second 0).AddSeconds((get-itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -name InstallDate).InstallDate)).ToLocalTime().AddHours((get-date -f zz))
	}
	
	if ( ($FeatureUpdateInstallDate.installDate -gt (Get-Date).AddDays(-$ThresholdUpdateDays)) -and ($NULL -ne $FeatureUpdateInstallDate) ){
		$FeatureUpdateDateDiff = [int][Math]::Ceiling((New-Timespan -Start $FeatureUpdateInstallDate.installDate -End (Get-Date) ).TotalDays)
		Write-Host ($StringFeatureUpdate -f $FeatureUpdateDateDiff)
		ScriptExit 0
	} else {
		Notify -Error $StringNoUpdatesFound
		ScriptExit 1004
	}
}

$updateHistory = $updateSearcher.QueryHistory(0, $updateHistoryCount)

[int]$UpdatesToInstallCount = 0
[string]$UpdatesToInstall = ""

[int]$FailedUpdatesCount = 0
[string]$FailedUpdates = ""

[int]$FailedUpdatesTotalCount = 0
[string]$FailedUpdatesTotal = ""

[int]$UpdatesWithinLast2Months = 0
[string]$InstalledUpdates = ""

[DateTime]$LastUpdateAt = (Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()
[string]$LastUpdate = ""

foreach ($Upd in $updateHistory) {
	if ((($Upd.operation -eq 1 -and $Upd.resultcode -eq 0) -or ($Upd.operation -eq 1 -and $Upd.resultcode -eq 1)) -and (($Upd.ClientApplicationID -eq "UpdateOrchestrator") -or ($Upd.ClientApplicationID -eq "AutomaticUpdates") -or ($Upd.ClientApplicationID -eq "AutomaticUpdatesWuApp"))) {
		$UpdatesToInstall += $Upd.Title + "`n"
		$UpdatesToInstallCount++
	}
	
	if ((($Upd.operation -eq 1 -and $Upd.resultcode -eq 4) -or ($Upd.operation -eq 1 -and $Upd.resultcode -eq 5)) -and (($Upd.ClientApplicationID -eq "UpdateOrchestrator") -or ($Upd.ClientApplicationID -eq "AutomaticUpdates") -or ($Upd.ClientApplicationID -eq "AutomaticUpdatesWuApp"))) {
		if (([DateTime]$Upd.Date) -gt (Get-Date).AddHours(-$ThresholdFailedUpdateHours)){
			$FailedUpdates += $Upd.Title + "`n"
			$FailedUpdatesCount++
		}
		$FailedUpdatesTotal += $Upd.Title + "`n"
		$FailedUpdatesTotalCount++
	}
	
	if (((($Upd.operation -eq 1 -and $Upd.resultcode -eq 2) -or ($Upd.operation -eq 1 -and $Upd.resultcode -eq 3)) -and (($Upd.ClientApplicationID -eq "UpdateOrchestrator") -or ($Upd.ClientApplicationID -eq "AutomaticUpdates") -or ($Upd.ClientApplicationID -eq "AutomaticUpdatesWuApp"))) -and ([DateTime]$Upd.Date) -gt (Get-Date).AddDays(-$ThresholdUpdateDays)) {
		$InstalledUpdates += ([DateTime]$Upd.Date).ToShortDateString() + " | " + $Upd.Title + "`n"
		$UpdatesWithinLast2Months++
	}
	
	if (((($Upd.operation -eq 1 -and $Upd.resultcode -eq 2) -or ($Upd.operation -eq 1 -and $Upd.resultcode -eq 3)) -and (($Upd.ClientApplicationID -eq "UpdateOrchestrator") -or ($Upd.ClientApplicationID -eq "AutomaticUpdates") -or ($Upd.ClientApplicationID -eq "AutomaticUpdatesWuApp"))) -and ([DateTime]$Upd.Date -gt $LastUpdateAt)){
		$LastUpdateAt = [DateTime]$Upd.Date
		$LastUpdate = $Upd.Title
	}
}

[string]$LastUpdateAtDate = $LastUpdateAt.ToLongDateString()

if ($LastUpdateAt -eq (Get-Date -Date "1970-01-01 00:00:00Z").ToUniversalTime()){
	Notify -Error $StringNoSuccessfulUpdate`n
	if ($UpdatesToInstallCount -gt 0){
		Notify ($StringUpdatesWaitingForInstallation -f $UpdatesToInstallCount, $UpdatesToInstall)
	}
	if ($FailedUpdatesTotalCount -gt 0){
		Notify ($StringErrorWhileInstallingUpdates -f $FailedUpdatesTotalCount, $FailedUpdatesTotal)
	}
	ScriptExit 1003
}elseif ( $FailedUpdatesCount -gt 0 -and !$NoCheckFailed ){
	Notify -Error ($StringErrorWhileInstallingUpdates -f $FailedUpdatesCount, $FailedUpdates)
	Notify ($StringLastUpdateInstalledon -f $LastUpdateAtDate, $LastUpdate)
	ScriptExit 1002
}elseif ( $UpdatesWithinLast2Months -le 0 ){
	Notify -Error ($StringNoUpdatesInstalledInTimePeriod) `n
	if ($UpdatesToInstallCount -gt 0){
		Notify ($StringUpdatesWaitingForInstallation -f $UpdatesToInstallCount, $UpdatesToInstall)
	}
	Notify ($StringLastUpdateInstalledon -f $LastUpdateAtDate, $LastUpdate)
	ScriptExit 1001
}else{
	Write-Host $StringNoUpdatesFailed `n
	Write-Host ($StringUpdatesInstalledinTimePeriod -f $UpdatesWithinLast2Months, $InstalledUpdates)
	if ($UpdatesToInstallCount -gt 0){
		Notify ($StringUpdatesWaitingForInstallation -f $UpdatesToInstallCount, $UpdatesToInstall)
	}
	ScriptExit 0
}
