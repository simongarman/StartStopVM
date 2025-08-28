<#
.SYNOPSIS
    Automates scheduled shutdown/startup of Azure VMs via "AutoShutdownSchedule" tag defines the time period the VM should be powered off using UTC

.DESCRIPTION
    Evaluates tagged schedules on VMs or resource groups, then enforces power state accordingly.
    Use "Simulate" mode to preview actions without executing them.
    
    Schedule Format Examples:
    - "10am > 6pm" (shutdown 10AM to 6PM)
    - "Saturday, Sunday, 10PM > 5AM" (shutdown weekends 10PM to 5AM)
    - "Monday, Tuesday, Wednesday, 9AM > 6PM" (shutdown weekdays 9AM-6PM)

.PARAMETER Simulate
    Boolean switch to enable simulation mode (no power actions performed).

.INPUTS
    None.

.OUTPUTS
    Informational and error messages for human review.
#>

# Outstanding Tasks:
# Error trapping and alerting 
# Send email alerts on error

param (
    [bool]$Simulate = $false
)

# =========================
# === Variables ===
# =========================
$keyVaultName = "kv-trg-uks-eval-02"
$secretName = "PSScriptSecret"
$tenantId = "0365b5b7-eb9d-4c7c-b42e-d4642c229292"
$clientId = "your-client-id-here"  # Add your actual client ID
$userId = 'it.helpdesk@riverside.org.uk'
$SubscriptionID = '009ccf42-7599-4656-b60a-0b0732be197d'
$recipientEmailrivtech = "simon.garman@riverside.org.uk"
$recipientEmailrivvmowner = "simon.garman@riverside.org.uk"
$VERSION = "2.1.2"

function Get-ClientSecretCredential {
    try {
        Connect-AzAccount -Identity | Out-Null
        
        $secureSecret = (Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $secretName).SecretValue
        $clientSecret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
        )

        $ClientSecretPass = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $ClientSecretPass

        return $credential
    }
    catch {
        Write-Output "ERROR: Failed to get client secret credential: $($_.Exception.Message)"
        throw
    }
}

function Connect-ToGraph {
    try {
        $credential = Get-ClientSecretCredential
        Connect-MgGraph -TenantId $tenantId -ClientSecretCredential $credential
    }
    catch {
        Write-Output "ERROR: Failed to connect to Graph: $($_.Exception.Message)"
        throw
    }
}

function Send-ErrorEmail {
    param (
        [string]$errorMessage
    )
    try {
        $subject = "Azure VM Automation Error Alert"
        $body = "An error occurred during the execution of the VM automation script:`n$errorMessage"
        $recipients = @($recipientEmailrivtech, $recipientEmailrivvmowner)
        Send-MailMessage -From $userId -To $recipients -Subject $subject -Body $body -SmtpServer "smtp.riverside.org.uk"
    }
    catch {
        Write-Output "WARNING: Failed to send error email: $($_.Exception.Message)"
    }
}

function CheckScheduleEntry {
    param([string]$TimeRange)

    try {
        $currentTime = (Get-Date).ToUniversalTime()
        $currentDay = $currentTime.DayOfWeek.ToString()
        
        # Debug output
        Write-Output "DEBUG: Checking schedule '$TimeRange' against current time $($currentTime.ToString('yyyy-MM-dd HH:mm:ss')) ($currentDay)"

        # Split by comma first to handle multiple day/time combinations
        $scheduleSegments = $TimeRange -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($segment in $scheduleSegments) {
            Write-Output "DEBUG: Processing segment: '$segment'"
            
            # Split each segment by spaces and commas to get individual parts
            $parts = $segment -split '[\s,]+' | Where-Object { $_ -ne '' }
            
            $daysInSegment = @()
            $timeRange = $null
            
            # Separate days from time ranges
            foreach ($part in $parts) {
                $part = $part.Trim().TrimEnd(',')
                Write-Output "DEBUG: Processing part: '$part'"
                
                if ($part -match '^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)$') {
                    $daysInSegment += $part
                    Write-Output "DEBUG: Found day: $part"
                }
                elseif ($part -match '(\d{1,2})(am|pm|AM|PM)\s*>\s*(\d{1,2})(am|pm|AM|PM)') {
                    $timeRange = $part
                    Write-Output "DEBUG: Found time range: $part"
                }
            }
            
            # Check if current day matches any day in this segment
            $dayMatch = $daysInSegment -contains $currentDay
            $timeMatch = $false
            
            if ($dayMatch -and $timeRange) {
                Write-Output "DEBUG: Day match found for segment. Checking time range: '$timeRange'"
                
                # Parse time range
                if ($timeRange -match '^(\d{1,2})(am|pm|AM|PM)\s*>\s*(\d{1,2})(am|pm|AM|PM)$') {
                    $times = $timeRange -split '>'
                    $startStr = $times[0].Trim()
                    $endStr = $times[1].Trim()
                    
                    Write-Output "DEBUG: Parsing time range: '$startStr' > '$endStr'"
                    
                    # Convert to 24-hour format for comparison
                    $startHour = [int]($startStr -replace '(?i)[ap]m', '')
                    $endHour = [int]($endStr -replace '(?i)[ap]m', '')
                    
                    # Convert AM/PM to 24-hour (case insensitive)
                    if ($startStr -match '(?i)pm' -and $startHour -ne 12) { $startHour += 12 }
                    if ($startStr -match '(?i)am' -and $startHour -eq 12) { $startHour = 0 }
                    if ($endStr -match '(?i)pm' -and $endHour -ne 12) { $endHour += 12 }
                    if ($endStr -match '(?i)am' -and $endHour -eq 12) { $endHour = 0 }
                    
                    $currentHour = $currentTime.Hour
                    
                    Write-Output "DEBUG: Time comparison - Current: $currentHour, Start: $startHour, End: $endHour"
                    
                    # Check if current time falls within the shutdown window
                    if ($startHour -gt $endHour) {
                        # Crosses midnight (e.g., 10PM -> 5AM)
                        if ($currentHour -ge $startHour -or $currentHour -le $endHour) {
                            $timeMatch = $true
                            Write-Output "DEBUG: Time match found (crosses midnight)"
                        }
                    } else {
                        # Same day (e.g., 10AM -> 6PM)
                        if ($currentHour -ge $startHour -and $currentHour -le $endHour) {
                            $timeMatch = $true
                            Write-Output "DEBUG: Time match found (same day)"
                        }
                    }
                }
            }
            
            # If any segment matches both day and time, return true
            if ($dayMatch -and $timeMatch) {
                Write-Output "DEBUG: Schedule match found in segment: '$segment'"
                return $true
            }
            
            Write-Output "DEBUG: Segment evaluation - Day match: $dayMatch, Time match: $timeMatch"
        }

        Write-Output "DEBUG: No schedule matches found"
        return $false
    } catch {
        Write-Output "WARNING: Failed to evaluate schedule '$TimeRange': $($_.Exception.Message)"
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

    try {
        if ($vm.ResourceType -eq "Microsoft.ClassicCompute/virtualMachines") {
            $classicVM = $classicVMList | Where-Object { $_.Name -eq $vm.Name }

            if ($DesiredState -eq "Started") {
                if ($classicVM.PowerState -notmatch "Started|Starting") {
                    if ($Simulate) {
                        Write-Output "[$($vm.Name)] SIMULATION -- Would start Classic VM"
                    } else {
                        Write-Output "[$($vm.Name)] Starting VM (outside shutdown schedule)"
                        $classicVM | Start-AzVM
                    }
                } else {
                    Write-Output "[$($vm.Name)] Classic VM already running"
                }
            }
            elseif ($DesiredState -eq "StoppedDeallocated") {
                if ($classicVM.PowerState -ne "Stopped") {
                    if ($Simulate) {
                        Write-Output "[$($vm.Name)] SIMULATION -- Would stop Classic VM"
                    } else {
                        Write-Output "[$($vm.Name)] Stopping VM (within shutdown schedule)"
                        $classicVM | Stop-AzVM -Force
                    }
                } else {
                    Write-Output "[$($vm.Name)] Classic VM already stopped"
                }
            }
        }
        elseif ($vm.ResourceType -eq "Microsoft.Compute/virtualMachines") {
            $status = (Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status).Statuses |
                      Where-Object { $_.Code -like "PowerState*" } |
                      Select-Object -First 1
            $state = $status.Code -replace "PowerState/", ""

            Write-Output "[$($vm.Name)] Current state: $state, Desired state: $DesiredState"

            if ($DesiredState -eq "Started" -and $state -ne "running") {
                if ($Simulate) {
                    Write-Output "[$($vm.Name)] SIMULATION -- Would start VM"
                } else {
                    Write-Output "[$($vm.Name)] Starting VM (outside shutdown schedule)"
                    Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name | Start-AzVM
                }
            }
            elseif ($DesiredState -eq "StoppedDeallocated" -and $state -ne "deallocated") {
                if ($Simulate) {
                    Write-Output "[$($vm.Name)] SIMULATION -- Would stop VM"
                } else {
                    Write-Output "[$($vm.Name)] Stopping VM (within shutdown schedule)"
                    Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name | Stop-AzVM -Force
                }
            }
            else {
                Write-Output "[$($vm.Name)] VM already in desired state"
            }
        }
        else {
            Write-Output "[$($vm.Name)] VM type not recognized: [$($vm.ResourceType)]. Skipping."
        }
    }
    catch {
        Write-Output "ERROR: Failed to assert power state for VM [$($vm.Name)]: $($_.Exception.Message)"
        Send-ErrorEmail -errorMessage "Failed to assert power state for VM [$($vm.Name)]: $($_.Exception.Message)"
    }
}

# =========================
# === Main Execution ===
# =========================

$currentTime = (Get-Date).ToUniversalTime()

Write-Output "========================================="
Write-Output "Azure VM Automation Script"
Write-Output "Version: $VERSION"
Write-Output "========================================="

if ($Simulate) {
    Write-Output "*** SIMULATION MODE: No power actions will be taken ***"
} else {
    Write-Output "*** LIVE MODE: Power actions will be enforced ***"
}

Write-Output "Current UTC time: [$($currentTime.ToString("yyyy-MM-dd HH:mm:ss"))] ($($currentTime.DayOfWeek))"

try {
    # Authenticate
    Write-Output "Authenticating to Azure..."
    Connect-AzAccount -Identity
    Select-AzSubscription -SubscriptionId $SubscriptionID
    Write-Output "Successfully authenticated to Azure"

    Write-Output "Retrieving VM information..."
    $resourceManagerVMList = Get-AzResource | Where-Object { $_.ResourceType -like "Microsoft.*/virtualMachines" } | Sort-Object Name
    $classicVMList = Get-AzVM
    $taggedGroups = Get-AzResourceGroup | Where-Object { $_.Tags["AutoShutdownSchedule"] }

    Write-Output "Found $($resourceManagerVMList.Count) VMs to evaluate"
    Write-Output "Found $($taggedGroups.Count) resource groups with AutoShutdownSchedule tags"

    $processedCount = 0
    $startedCount = 0
    $stoppedCount = 0
    $skippedCount = 0

    foreach ($vm in $resourceManagerVMList) {
        $processedCount++
        Write-Output ""
        Write-Output "Processing VM [$($vm.Name)] ($processedCount of $($resourceManagerVMList.Count))"
        