<#
.DESCRIPTION
    Set Teams firewall rules, now with plenty try & catch

.EXAMPLE
    Deploy as a proactive remediation, see status in Intune
    
.NOTES
    Created: 2023-01-26
    Updated: 2023-02-16
    Author: returaxel
    ToDo: fancy-up try/catch
#>

$MeasureCommand = Measure-Command { # START MEASURE

# Current User
try {
    $CurrentUser = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object UserName).UserName.Split('\')[1]
    }
catch {
    Write-Output "CurrentUser is NULL: $($PSItem.Exception)"
    exit 1
}

# Get local path for found user
$UserProfile = (Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special = 0' | Where-Object -Property LocalPath -match $CurrentUser).LocalPath

# We've seen duplicate paths in LocalPath, this should skip any profile.domain 
$UserProfile.Split(' ') | ForEach-Object {
        if (([regex]::Match($_,"$CurrentUser$")).Success) {
        [string] $LocalPath = $_
    }
}

if ([string]::IsNullOrEmpty($LocalPath)) {
    Write-Output "LocalPath is NULL."
    exit 1
}

# Concatinate path to Teams.exe 
[string] $TeamsPath ='{0}{1}' -f $LocalPath, "\AppData\Local\Microsoft\Teams\Current\Teams.exe"

# Remove old fw rules, if any
try {
    Get-NetFirewallApplicationFilter -Program $TeamsPath -ErrorAction Stop | Remove-NetFirewallRule -ErrorAction Stop
}
catch {
    Write-Output "Error removing rules:`n $($PSItem.Exception)"
} 

# Add new fw rules
$RuleName = "Teams.exe, $($CurrentUser)"
try {
    New-NetFirewallRule -DisplayName $RuleName -Direction Inbound -Profile Domain -Program $TeamsPath -Action Allow -Protocol Any -ErrorAction Stop | Out-Null
    New-NetFirewallRule -DisplayName $RuleName -Direction Inbound -Profile Public,Private -Program $TeamsPath -Action Block -Protocol Any -ErrorAction Stop | Out-Null
    }
catch {
    Write-Output "Error adding rules:`n $($PSItem.Exception)"
    $Error.Clear()
    exit 1
}

} # END MEASURE

# For logging purposes, this will be in Intune if rules were added successfully
Write-Output "$($CurrentUser), $($LocalPath), $($MeasureCommand.Milliseconds)ms, $([datetime]::Now)"
exit 0
