workflow startstopvm
{

# Subscriptions excluded from this runbook execution
$excludedSubs = @((Get-AutomationVariable -Name 'excludedSubs').split(","))

# Ensures you do not inherit an AzContext in your runbook
$disableContext = Disable-AzContextAutosave -Scope Process

# Connect to Azure with user-assigned managed identity
$null = Connect-AzAccount -Identity -AccountId REPLACE_WITH_USER_MANAGED_IDENTITY_ID

# Gets the list of all subscriptions
$subscriptions = Get-AzSubscription | ? SubscriptionId -notin $excludedSubs | foreach { $_.SubscriptionId }

#  Get day of the week
$day = Get-DayOfWeek

# Query to get all VMs which are enable for start stop. It will only get VMs which are scheduled for today.
$query = "resources | where type in~ ('microsoft.compute/virtualmachines') | where tags['vm-start-stop-enable'] == 'true' | where tags['vm-start-stop-schedule'] contains '$day' or (tags['vm-start-stop-schedule'] contains 'weekends' and 'sat-sun' contains '$day') or (tags['vm-start-stop-schedule'] contains 'weekdays' and 'mon-tue-wed-thu-fri' contains '$day')"

# Execute Query
$vms = Search-AzGraph -Subscription $subscriptions -Query $query

Write-Output "Query matched $($vms.count) VMs."

# Group VMs in the same Resource Group
$groupVMs = $vms | Group-Object resourceGroup

Write-Output "VMs grouped in the following RGs."
Write-Output $groupVMs.Name

function Start-Validation {
	param (
		$vm
	)
	$output = @{}
	$notValid = $false
	if(($vm.tags.'vm-start-stop-sequence' -ne $null) -and ($vm.tags.'vm-start-stop-sequence' -notmatch "\d{1}")){
		$message = 'vm-start-stop-sequence must be integer.'
		$notValid = $true
		$output = $output + @{$notValid=$message}
	}
	if(!($vm.tags.'vm-start-stop-schedule') -and $output.count -eq 0){
		$message = 'vm-start-stop-schedule tag must exist and must not be empty.'
		$notValid = $true
		$output = $output + @{$notValid=$message}
	}
	if($vm.tags.'vm-start-stop-schedule' -match "\s+" -and $output.count -eq 0){
		$message = 'vm-start-stop-schedule must not have White spaces.'
		$notValid = $true
		$output = $output + @{$notValid=$message}
	}
	if($vm.tags.'vm-start-stop-schedule' -notmatch '^(mon|tue|wed|thu|fri|sat|sun|weekends|weekdays)=[SE]\d{1,2}:\d{2}' -and $output.count -eq 0){
		$message = 'vm-start-stop-schedule is not in the required format.'
		$notValid = $true
		$output = $output + @{$notValid=$message}
	}
	if($notValid){return $output}
	else{return $true}
}

function Get-DayOfWeek {

	# Get current UTC Date Time
	$currentDate = (Get-Date).ToUniversalTime()
	# Get day of the week
	$dw = $currentDate.DayOfWeek.ToString().ToLower().SubString(0, 3)

	return $dw
    
}

function Get-StartStopTag {
    param (
		$vm
    )
	
	$currentDate = (Get-Date).ToUniversalTime()
	$dw = $currentDate.DayOfWeek.ToString().ToLower().SubString(0, 3)
	$days = $vm.tags."vm-start-stop-schedule"
	$arrayOfDays = @($days.split("|")).replace("weekdays","mon|tue|wed|thu|fri").replace("weekends","sat|sun")
	$arrayOfDays = [array]$arrayOfDays -match $dw
	# In case weekends or weekdays are specified in addition to overlapping day of the week, select the shortest string which will be the day of the week which is more specific 
	$arrayOfDays = $arrayOfDays | Sort-Object length | Select-Object -First 1

	return $arrayOfDays
    
}

# Iterate through each group of VMs in the same RG in parallel
foreach -Parallel ($group in $groupVMs) {
	# If there is 1 or less VMs which have the sequence tag in the same RG then sequence is not required
	if(($group.group.tags -match "vm-start-stop-sequence").count -le 1){
		# Iterate through each VM in parallel
		foreach -Parallel ($vm in $group.Group){
			$valid = Start-Validation -vm $vm
			if($valid -eq $true){
				$currentDate = (Get-Date).ToUniversalTime()
				$time = $currentDate.TimeOfDay.TotalMinutes
				# Get VM schedule
				$arrayOfDays = Get-StartStopTag -vm $vm
				# Get just the time(s)
				$schedules = ($arrayOfDays.split("=")[1].split("-"))
				# If there is only a single time and it starts with S, then that is the Start Time
				if($schedules.count -eq 1 -and $schedules -match "S"){
					Write-Output "$($vm.Name) set to start at $utcStartTime"
					$utcStartTime = [DateTime]$schedules.replace("S","")
					$singleTime  = $true
				}
				# If there is only a single time and it starts with E, then that is the Stop Time
				if($schedules.count -eq 1 -and $schedules -match "E"){
					$utcStopTime = [DateTime]$schedules.replace("E","")
					Write-Output "$($vm.Name) set to stop at $utcStopTime"
					$singleTime  = $true
				}
				# If there are two times, then the first is the Start Time and the second is the Stop Time
				if($schedules.count -eq 2){
					Write-Output "$($vm.Name) set to start at $utcStartTime"
					Write-Output "$($vm.Name) set to stop at $utcStopTime"
					$utcStartTime = [DateTime]$schedules[0].replace("S","")
					$utcStopTime = [DateTime]$schedules[1].replace("E","")
				}
				# Transform the time in Total Minutes
				$utcStartTimeTotalMinutes = $utcStartTime.TimeOfDay.TotalMinutes
				$utcStopTimeTotalMinutes = $utcStopTime.TimeOfDay.TotalMinutes
				# Work out duration of downtime
				if($singleTime -ne $true){
					if(($utcStartTime-$utcStopTime).TotalHours -is [int]){
						$duration = ($utcStartTime-$utcStopTime).TotalHours
					}
					else {
						$duration = ($utcStartTime-$utcStopTime).TotalHours + 24
					}
				}
				# If current time is greater or equal the (time to start - 15 minutes) and current time is less or equal the (time to start + 15 minutes) and VM is not running or starting
				# This means the VM may start 15 minutes earlier but, in theory, never later than the schedule 
				if ($time -ge ($utcStartTimeTotalMinutes - 15) -and $time -le ($utcStartTimeTotalMinutes + 15) -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "running" -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "starting") {
					# Select VM subscription
					$currentSub = Select-AzSubscription -SubscriptionId $vm.SubscriptionId
					Write-Output "Starting VM $($vm.Name) at $(Get-Date)..."
					Start-AzVM -Name $vm.Name -ResourceGroupName $vm.resourceGroup -NoWait
				}
				# If current time is greater or equal the time to stop and current time is less than the (time to start + 15 minutes) and VM is not deallocated or stopping
				# This means the VM may stop 15 minutes later, but never earlier than the schedule 
				if ($time -ge $utcStopTimeTotalMinutes -and $time -lt ($utcStopTimeTotalMinutes + 15) -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "deallocated" -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "deallocating") {
					# Select VM subscription
					$currentSub = Select-AzSubscription -SubscriptionId $vm.SubscriptionId
					Write-Output "Stopping VM $($vm.Name) at $(Get-Date)..."
					# Stop VM
					Stop-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroup -Force -NoWait		
				}
			}
			else{
				Write-Output "$($vm.name): $($valid.Values)"
			}
		}
	}
	else{
		$arrayofVMsToStart = @()
		$arrayofVMsToStop = @()
		foreach($vm in $group.Group){
			$valid = Start-Validation -vm $vm
			if($valid -eq $true){
				$currentDate = (Get-Date).ToUniversalTime()
				$time = $currentDate.TimeOfDay.TotalMinutes
				$dw = Get-DayOfWeek
				# Get schedules
				$arrayOfDays = Get-StartStopTag -vm $vm
				# Get just the time(s)
				$schedules = ($arrayOfDays.split("=")[1].split("-"))
				# If there is only a single time and it starts with S, then that is the Start Time
				if($schedules.count -eq 1 -and $schedules -match "S"){
					$utcStartTime = ([DateTime]$schedules.replace("S","")).ToUniversalTime()
					Write-Output "$($vm.Name) set to start at $utcStartTime"
					$singleTime  = $true
				}
				# If there is only a single time and it starts with E, then that is the Stop Time
				if($schedules.count -eq 1 -and $schedules -match "E"){
					$utcStopTime = ([DateTime]$schedules.replace("E","")).ToUniversalTime()
					Write-Output "$($vm.Name) set to stop at $utcStopTime"
					$singleTime  = $true
				}
				# If there are two times, then the first is the Start Time and the second is the Stop Time
				if($schedules.count -eq 2){
					$utcStartTime = ([DateTime]$schedules[0].replace("S","")).ToUniversalTime()
					$utcStopTime = ([DateTime]$schedules[1].replace("E","")).ToUniversalTime()
					Write-Output "$($vm.Name) set to start at $utcStartTime"
					Write-Output "$($vm.Name) set to stop at $utcStopTime"
				}
				# Transform the time in Total Minutes
				$utcStartTimeTotalMinutes = $utcStartTime.TimeOfDay.TotalMinutes
				$utcStopTimeTotalMinutes = $utcStopTime.TimeOfDay.TotalMinutes
				if($singleTime -ne $true){
					# Work out duration of downtime
					if(($utcStartTime-$utcStopTime).TotalHours -is [int]){
						$duration = ($utcStartTime-$utcStopTime).TotalHours
					}
					else {
						$duration = ($utcStartTime-$utcStopTime).TotalHours + 24
					}
				}
				Write-Output "Current time in TotalMinutes is: $time"
				Write-Output "Start time in TotalMinutes is: $utcStartTimeTotalMinutes"
				Write-Output "Stop time in TotalMinutes is: $utcStopTimeTotalMinutes"
				
				Write-Output "$($vm.Name) downtime duration will be $duration"
				# If current time is greater or equal the (time to start - 15 minutes) and current time is less or equal the (time to start + 15 minutes) and VM is not running or starting
				# This means the VM may start 15 minutes earlier but, in theory, never later than the schedule 
				if ($time -ge ($utcStartTimeTotalMinutes - 15) -and $time -le ($utcStartTimeTotalMinutes + 15) -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "running" -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "starting") {
					# Select VM subscription
					$currentSub = Select-AzSubscription -SubscriptionId $vm.SubscriptionId
					# If VM should start in sequence
					Write-Output "Found VM to Start."
					if ($vm.tags."vm-start-stop-sequence") {
						# Create array of VMs to control sequence
						[array]$arrayofVMsToStart += $vm
						Write-Output "$($vm.Name) will be started in the sequence: $($vm.tags.'vm-start-stop-sequence')"
					}
					# Else, just start the VM
					else {
						Write-Output "Starting VM $($vm.Name) at $(Get-Date)..."
						Start-AzVM -Name $vm.Name -ResourceGroupName $vm.resourceGroup -NoWait
						Start-Sleep 15
					}
				}
				# If current time is greater or equal the time to stop and current time is less than the (time to start + 15 minutes) and VM is not deallocated or stopping
				# This means the VM may stop 15 minutes later, but never earlier than the schedule 
				if ($time -ge $utcStopTimeTotalMinutes -and $time -lt ($utcStopTimeTotalMinutes + 15) -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "deallocated" -and $vm.properties.extended.instanceView.powerState.displayStatus -notmatch "deallocating") {
					# Select VM subscription
					$currentSub = Select-AzSubscription -SubscriptionId $vm.SubscriptionId
					Write-Output "Found VM to Stop."
					# If VM should stop in sequence
					if ($vm.tags."vm-start-stop-sequence") {
						# Create array of VMs to control sequence
						[array]$arrayofVMsToStop += $vm
						Write-Output "$($vm.Name) will be stopped in the sequence: $($vm.tags.'vm-start-stop-sequence')"
					}
					# Else, just stop the VM
					else {
						Write-Output "Stopping VM $($vm.Name) at $(Get-Date)..."
						# Stop VM
						Stop-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroup -Force -NoWait
					}
				}
			}
			else{
				Write-Output "$($vm.name): $($valid.Values)"
			}
		}

		Write-Output "Count of VMs to stop in sequence is: $($arrayofVMsToStop.count)"
		Write-Output "Count of VMs to start in sequence is: $($arrayofVMsToStart.count)"

		# Start VMs in sequence
		foreach($vm in $arrayofVMsToStart | Sort-Object -property @{e={$_.tags.'vm-start-stop-sequence'}}){
			Write-Output "Starting VM $($vm.Name) - Sequence $($_.tags.'vm-start-stop-sequence') at $(Get-Date)..."
			Start-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroup -NoWait
			Start-Sleep 15
		}
		# Stop VMs in sequence
		foreach($vm in ($arrayofVMsToStop | Sort-Object -property @{e={$_.tags.'vm-start-stop-sequence'}} -Descending)){
			Write-Output "Stopping VM $($vm.Name) at $(Get-Date)..."
			# Stop VM
			Stop-AzVM -Name $vm.Name -ResourceGroupName $vm.ResourceGroup -Force -NoWait
			Start-Sleep 15
		}
	}
}
}