try{
	$session = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$ComputerName))
	$updateSearcher = $session.CreateUpdateSearcher()
	$updateHistoryCount = $updateSearcher.GetTotalHistoryCount()
}catch [Exception]{
		Write-Host "Cannot connect to Windows Update service"
		Write-Host $_.Exception.Message
		$host.SetShouldExit(1005) 
		exit 1005
}

if ( $updateHistoryCount -le 0 ){

	[Nullable[datetime]]$FeatureUpdateInstallDate 
	Try{
		[datetime]$FeatureUpdateInstallDate = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop | Select-Object InstallDate).InstallDate
	}Catch [Exception]{
		[datetime]$FeatureUpdateInstallDate = ((get-date -year 1970 -month 1 -day 1 -hour 0 -minute 0 -second 0).AddSeconds((get-itemproperty -path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -name InstallDate).InstallDate)).ToLocalTime().AddHours((get-date -f zz))
	}
	
	if ( ($FeatureUpdateInstallDate.installDate -gt (Get-Date).AddDays(-45)) -and ($FeatureUpdateInstallDate -ne $NULL) ){
		$FeatureUpdateDateDiff = [int][Math]::Ceiling((New-Timespan -Start $FeatureUpdateInstallDate.installDate -End (Get-Date) ).TotalDays)
		Write-Host "Last feature update was $FeatureUpdateDateDiff days ago."
		$host.SetShouldExit(0) 
		exit 0
	} else {
		Write-Host "No updates found in update history."
		$host.SetShouldExit(1004) 
		exit 1005
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
		if (([DateTime]$Upd.Date) -gt (Get-Date).AddHours(-23)){
			$FailedUpdates += $Upd.Title + "`n"
			$FailedUpdatesCount++
		}
		$FailedUpdatesTotal += $Upd.Title + "`n"
		$FailedUpdatesTotalCount++
    }
	
	if (((($Upd.operation -eq 1 -and $Upd.resultcode -eq 2) -or ($Upd.operation -eq 1 -and $Upd.resultcode -eq 3)) -and (($Upd.ClientApplicationID -eq "UpdateOrchestrator") -or ($Upd.ClientApplicationID -eq "AutomaticUpdates") -or ($Upd.ClientApplicationID -eq "AutomaticUpdatesWuApp"))) -and ([DateTime]$Upd.Date) -gt (Get-Date).AddDays(-45)) {
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
	Write-Host "No successful update found in update history.`n"
	if ($UpdatesToInstallCount -gt 0){
		Write-Host "$UpdatesToInstallCount waiting for installation/reboot:`n$UpdatesToInstall"
	}
	if ($FailedUpdatesTotalCount -gt 0){
		Write-Host "Error while installing $FailedUpdatesTotalCount updates:`n$FailedUpdatesTotal"
	}
	$host.SetShouldExit(1003) 
    exit 1003
}elseif ( $FailedUpdatesCount -gt 0 ){
    Write-Host "Error while installing  $FailedUpdatesCount updates:`n$FailedUpdates"
	Write-Host "Last update installed on $LastUpdateAtDate ($LastUpdate)"
	$host.SetShouldExit(1002) 
    exit 1002
}elseif ( $UpdatesWithinLast2Months -le 0 ){
    Write-Host "No updates have been installed within the last 45 days`n"
	if ($UpdatesToInstallCount -gt 0){
		Write-Host "$UpdatesToInstallCount waiting for installation/reboot:`n$UpdatesToInstall"
	}
	Write-Host "Last update installed on $LastUpdateAtDate ($LastUpdate)"
	$host.SetShouldExit(1001) 
	exit 1001
}else{
    Write-Host "No updates failed to install.`n"
	Write-Host "$UpdatesWithinLast2Months installed updates in the last 45 days:`n$InstalledUpdates"
	if ($UpdatesToInstallCount -gt 0){
		Write-Host "$UpdatesToInstallCount waiting for installation/reboot:`n$UpdatesToInstall"
	}
	$host.SetShouldExit(0) 
	exit 0
}

# https://docs.microsoft.com/en-us/windows/desktop/api/wuapi/
# https://docs.microsoft.com/en-us/previous-versions/windows/desktop/aa386400(v%3dvs.85)
