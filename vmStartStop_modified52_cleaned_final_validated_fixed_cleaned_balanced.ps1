<#
.SYNOPSIS
    Automates scheduled shutdown/startup of Azure VMs via "AutoShutdownSchedule" tag defines the time period the VM should be powered off using UTC

.DESCRIPTION
    Evaluates tagged schedules on VMs or resource groups, then enforces power state accordingly.
    Use "Simulate" mode to preview actions without executing them.

.PARAMETER Simulate
    Boolean switch to enable simulation mode (no power actions performed).

.INPUTS
    None.

.OUTPUTS
    Informational and error messages for human review.
#>


# Oustanding Tasks:
# Error trapping and alerting 
# Send email alerts on error

param (
    [bool]$Simulate = $false
)


function Get-ClientSecretCredential {
    Connect-AzAccount -Identity | Out-Null

Connect-ToGraph
    $secureSecret = (Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName).SecretValue
    $clientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
    )

    $ClientSecretPass = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $ClientSecretPass

    return $credential
}
function Connect-ToGraph {
    $credential = Get-ClientSecretCredential
    Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential
}

# =========================
# === Variables ===
# =========================


$keyVaultName = "kv-trg-uks-eval-02"
$secretName = "PSScriptSecret"
$tenantId = "0365b5b7-eb9d-4c7c-b42e-d4642c229292"
$userId = 'it.helpdesk@riverside.org.uk'
$SubscriptionID = '009ccf42-7599-4656-b60a-0b0732be197d'
$recipientEmailrivtech = "simon.garman@riverside.org.uk"
$recipientEmailrivvmowner = "simon.garman@riverside.org.uk"




function Send-ErrorEmail {
    param (
        [string]$errorMessage
    )
    $subject = "Azure VM Automation Error Alert"
    $body = "An error occurred during the execution of the VM automation script:`n$errorMessage"
    $recipients = "$recipientEmailrivtech,$recipientEmailrivvmowner"
    Send-MailMessage -From $userId -To $recipients -Subject $subject -Body $body -SmtpServer "smtp.riverside.org.uk"
}

$VERSION = "2.1.1"
$currentTime = (Get-Date).ToUniversalTime()

Write-Output "Runbook started. Version: $VERSION"

if ($Simulate) {
    Write-Output "*** SIMULATION MODE: No power actions will be taken ***"
} else {
    Write-Output "*** LIVE MODE: Power actions will be enforced ***"
}

Write-Output "Current UTC time: [$($currentTime.ToString("yyyy-MM-dd HH:mm:ss"))]"

# Authenticate
Connect-AzAccount -Identity
Select-AzSubscription -SubscriptionId $SubscriptionID

function CheckScheduleEntry {
    param([string]$TimeRange)

    try {
        $currentTime = (Get-Date).ToUniversalTime()
        $currentDay = $currentTime.DayOfWeek.ToString()
        $now = [datetime]::ParseExact($currentTime.ToString('HHmm'), 'HHmm', $null)

        $ScheduleParts = $TimeRange -split '\s+'
        $dayMatch = $false
        $timeMatch = $false

        foreach ($part in $ScheduleParts) {
            # Check for day match
            if ($part -match '^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$') {
                if ($part -eq $currentDay) {
                    $dayMatch = $true
                }
            }
            # Check for time range match
            elseif ($part -match '^\d{1,2}(AM|PM)\s*->\s*\d{1,2}(AM|PM)$') {
                $times = $part -split '->'
                $start = [datetime]::ParseExact($times[0].Trim(), 'htt', $null)
                $end = [datetime]::ParseExact($times[1].Trim(), 'htt', $null)

                if ($start -gt $end) {
                    if ($now -ge $start -or $now -le $end) {
                        $timeMatch = $true
                    }
                } else {
                    if ($now -ge $start -and $now -le $end) {
                        $timeMatch = $true
                    }
                }
            }
        }

        return ($dayMatch -and $timeMatch)
    } catch {
        Write-Output "WARNING: Failed to evaluate schedule: $($_.Exception.Message)"
        Send-ErrorEmail -errorMessage $_.Exception.Message
        return $false
    }
}

            } elseif ($part -match '^(\d{1,2}(AM|PM))\s*->\s*(\d{1,2}(AM|PM))$') {
                $times = $part -split '->'
                $start = [datetime]::ParseExact($times[0].Trim(), 'htt', $null)
                $end = [datetime]::ParseExact($times[1].Trim(), 'htt', $null)
                $now = [datetime]::ParseExact($currentTime.ToString('HHmm'), 'HHmm', $null)
                if ($start -gt $end) {
                    if ($now -ge $start -or $now -le $end) {
                        $timeMatch = $true
                    }
                } else {
                    if ($now -ge $start -and $now -le $end) {
                        $timeMatch = $true
                    }
                }
            }
        }
        return ($dayMatch -and $timeMatch)
    } catch {
        Write-Output "WARNING: Failed to evaluate schedule: $($_.Exception.Message)"
        Send-ErrorEmail -errorMessage $_.Exception.Message
        return $false
    }
}
function AssertPowerState {
    param (
        $vm,
        [string]$DesiredState,
        $resourceManagerVMList,
        $classicVMList,
        [bool]$Simulate
    )

    if ($vm.ResourceType -eq "Microsoft.ClassicCompute/virtualMachines") {
        $classicVM = $classicVMList | Where-Object { $_.Name -eq $vm.Name }

        if ($DesiredState -eq "Started") {
            if ($classicVM.PowerState -notmatch "Started|Starting") {
                if ($Simulate) {
                    Write-Output "[$($vm.Name)] SIMULATION -- Would start Classic VM"
                } else {
Write-Output "[$($vm.Name)] Starting VM due to schedule mismatch."
                    $classicVM | Start-AzVM
                }
            }
        }
        elseif ($DesiredState -eq "StoppedDeallocated") {
            if ($classicVM.PowerState -ne "Stopped") {
                if ($Simulate) {
                    Write-Output "[$($vm.Name)] SIMULATION -- Would stop Classic VM"
                } else {
Write-Output "[$($vm.Name)] Stopping VM due to schedule match."
                    $classicVM | Stop-AzVM -Force
                }
            }
        }
    }
    elseif ($vm.ResourceType -eq "Microsoft.Compute/virtualMachines") {
        $status = (Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status).Statuses |
                  Where-Object { $_.Code -like "PowerState*" } |
                  Select-Object -First 1
        $state = $status.Code -replace "PowerState/", ""

        if ($DesiredState -eq "Started" -and $state -ne "running") {
            if ($Simulate) {
                Write-Output "[$($vm.Name)] SIMULATION -- Would start VM"
            } else {
Write-Output "[$($vm.Name)] Starting VM due to schedule mismatch."
Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name | Start-AzVM
            }
        }
        elseif ($DesiredState -eq "StoppedDeallocated" -and $state -ne "deallocated") {
            if ($Simulate) {
                Write-Output "[$($vm.Name)] SIMULATION -- Would stop VM"
            } else {
Write-Output "[$($vm.Name)] Stopping VM due to schedule match."
Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name | Stop-AzVM -Force
            }
        }
    }
    else {
        Write-Output "[$($vm.Name)] VM type not recognized: [$($vm.ResourceType)]. Skipping."
    }
}

try {
    $resourceManagerVMList = Get-AzResource | Where-Object { $_.ResourceType -like "Microsoft.*/virtualMachines" } | Sort-Object Name
    $classicVMList = Get-AzVM
    $taggedGroups = Get-AzResourceGroup | Where-Object { $_.Tags["AutoShutdownSchedule"] }

    foreach ($vm in $resourceManagerVMList) {
        $schedule = $null

        if ($vm.Tags["AutoShutdownSchedule"]) {
            $schedule = $vm.Tags["AutoShutdownSchedule"]
        }
        elseif ($taggedGroups | Where-Object { $_.ResourceGroupName -eq $vm.ResourceGroupName }) {
            $group = $taggedGroups | Where-Object { $_.ResourceGroupName -eq $vm.ResourceGroupName }
            $schedule = $group.Tags["AutoShutdownSchedule"]
        }

        if (-not $schedule) {
            Write-Output "[$($vm.Name)] No schedule tag found. Skipping."
            continue
        }

        $ranges = $schedule -split "," | ForEach-Object { $_.Trim() }
        $matched = $false

        foreach ($r in $ranges) {
            if (CheckScheduleEntry -TimeRange $r) {
                $matched = $true
                break
            }
        }

        if ($matched) {
            $desiredState = "StoppedDeallocated"
        } else {
            $desiredState = "Started"
        }
        AssertPowerState -vm $vm -DesiredState $desiredState -resourceManagerVMList $resourceManagerVMList -classicVMList $classicVMList -Simulate $Simulate
catch {
    throw "Unexpected error: $($_.Exception.Message)"
    Send-ErrorEmail -errorMessage $_.Exception.Message
finally {
    $endTime = (Get-Date).ToUniversalTime()
    $duration = $endTime - $currentTime
    Write-Output "Runbook completed in $($duration.ToString("hh\:mm\:ss"))"
     Disconnect-AzAccount