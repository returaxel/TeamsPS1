<#
.DESCRIPTION
    Install, update and start Teams
    Download company background images
    CTRL + F "editme"

.EXAMPLE
    Deploy as a scheduled task, run when device is inactive for x minutes
    
.NOTES
    Created: 2021-03-xx
    Updated: 2023-02-16 (corrected things, thnx SkyN9ne)
    Author: returaxel
    ToDo: error handling
#>

# Hide progress, save time
$ProgressPreference = 'SilentlyContinue'

# Log
$logDir = "$($env:TEMP)\TeamsPS1\"
$logTxt = '{0}{1}' -F $logDir,"teamsps1.txt"

### ---------- Installation check
function Get-TeamsInstallationCheck { 

    [CmdletBinding()]
    param (
        [Parameter()][string]$teamsReg = "HKCU:\\SOFTWARE\Microsoft\Office\Teams",     
        [Parameter()][string]$teamsDed = "$($env:LOCALAPPDATA)\Microsoft\Teams\.dead", 
        [Parameter()][string]$teamsSet = "$($env:APPDATA)\Microsoft\Teams\settings.json", 
        [Parameter()][string]$teamsUpd = "$($env:LOCALAPPDATA)\Microsoft\Teams\update.exe",
        [Parameter()][switch]$checkSwi
    )

    Write-Host "`n[info] Status:"
    
    if ($checkSwi) {  
        if ((Test-Path $teamsDed) -or ((Get-ItemProperty $teamsReg).LoggedInOnce -ne 1)) {
            Write-Host "`t$($teamsDed)" -ForegroundColor Red  
            return Start-TeamsInstallation
        }
    }

    if ((Test-Path $teamsSet)) {
        $teamsVer = (Get-Content $teamsSet | ConvertFrom-Json).Version
        Write-Host "`tInstalled: $($teamsVer)" -ForegroundColor Green 
    } else {
        Write-Host "`tInstallation failed..." -ForegroundColor Red 
    }
}

# Update teams
function Start-TeamsUpdate {

    [CmdletBinding()]
    param (
        [Parameter()][datetime]$timeNow = [datetime]::Now,
        [Parameter()][string]$teamsUpd = "$($env:LOCALAPPDATA)\Microsoft\Teams\update.exe"
    )

    $timeLAT = (Get-Item $teamsUpd).LastAccessTime
    $timeSpan = (New-TimeSpan -Start $timeLAT -End $timeNow).Hours

    Write-Host "`n[info] Update:`n`tCheck for updates if (update.exe).accesstime < 72 hours`n`t Last access: $($timeSpan) hours ago"

    if ($timeSpan -ge 72) { # editme greater number to run less often
        Write-Host "`n[info] Update: Checking for updates..." -ForegroundColor Yellow
        Start-Process -FilePath $teamsUpd -ArgumentList "-s","--processStart teams.exe" -PassThru
        Start-Sleep 3
    }
}

### ---------- Installation
function Start-TeamsInstallation {

    [CmdletBinding()]
    param (
        [Parameter()][string]$teamsUrl = "https://go.microsoft.com/fwlink/p/?LinkID=869426&culture=en-us&country=WW&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64",
        [Parameter()][string]$teamsUpd = "$($env:LOCALAPPDATA)\Microsoft\Teams\update.exe",
        [Parameter()][string]$teamsExe = "$($env:TEMP)\TeamsPS1\Teams_windows_x64.exe" # .EXE saved here
    )

    if (!(Test-Path $teamsExe)) {
        Write-Host "`tDownloading..."
        Invoke-WebRequest -Uri $teamsUrl -OutFile $teamsExe -UseBasicParsing
    } 

    Write-Host "`tInstalling..."

    # Install silently
    Start-Process -FilePath $teamsExe -ArgumentList "-s" -Wait

    # Start
    Start-Process -FilePath $teamsUpd -ArgumentList "--processStart teams.exe"

    # Wait for teams process to start
    while (!$pcsAge) {
        $teamsPcs = Get-Process teams -ErrorAction SilentlyContinue

        if ($teamsPcs) {
            $pcsAge = $teamsPcs[0].StartTime.AddSeconds(8) -lt [datetime]::Now 
        }

        Start-Sleep 4
    }
    Get-TeamsInstallationCheck
}

### ---------- Copy background images
function Get-TeamsPS1Images {

    [CmdletBinding()]
    param (
        [Parameter()][string]$copySrc = "editme",
        [Parameter()][string]$copyDst = "editme",
        [Parameter()][int]$imgCount = 0
    )
    
    Write-Host "`n[info] Images:"

    if (Test-Path $copySrc) {

        if (!(Test-Path $copyDst)) {
            mkdir $copyDst
        }

        $bgRef = Get-ChildItem $copySrc -Name 
        $bgDiff = Get-ChildItem $copyDst -Exclude *thumb* -Name 

        foreach ($img in $bgRef) {
            if ($img -notin $bgDiff) {
                (Copy-Item -path $copySrc$img -Destination $copyDst -PassThru).Name  | Out-Host
                $imgCount += 1
            }
        }
        return Write-Host "`tImages updated: $($imgCount)"
    }
     Write-Host "`tNetwork unavailable." -ForegroundColor Yellow
}

$timed =  Measure-Command {

# Create directory for files and transcript
if (!(Test-Path $logDir)) {
    $null = mkdir $logDir
}

# Start logging
Start-Transcript -Path $logTxt -Force

# Check if Teams is installed, install if not
Get-TeamsInstallationCheck -checkSwi

# Run update.exe if lastaccesstime < 72 hours
Start-TeamsUpdate 

# Only copy images if fileshare is accessable
Get-TeamsPS1Images 

# Activate Teams Add-in for Outlook 
$addinPath = "HKCU:\\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
$addinCheck = (Get-ItemProperty -Path $addinPath -ErrorAction SilentlyContinue).LoadBehavior -ne 3

if ($addinCheck) {
    Write-Host "`n[info] Add-in:`n`tEnabled Teams Add-in for Outlook [ LoadBehavior = 3 ]"
    Set-ItemProperty -Path $addinPath -Name LoadBehavior -Value 3
}

} # END Measure command

Write-Host "`n[info] Runtime:" 
Write-Host "`t$($timed.TotalSeconds) seconds" -ForegroundColor Cyan

# End logging
Stop-Transcript
