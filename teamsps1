# Hide progress, save time
$ProgressPreference = 'SilentlyContinue'

# Log
$logDir = "$($env:TEMP)\TeamsPS1\"
$logTxt = '{0}{1}' -F $logDir,"teamsp1.txt"

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
        $dedCheck = Test-Path $teamsDed
        $regCheck = (Get-ItemProperty $teamsReg).LoggedInOnce -eq 1
    
        if ($dedCheck -or !($regCheck)) {
            Write-Host "`t$($teamsDed)" -ForegroundColor Red  
            return Start-TeamsInstallation
        }
    }

    $insCheck = Test-Path $teamsSet

    if ($insCheck) {
        $teamsVer = (Get-Content $teamsSet | ConvertFrom-Json).Version
        Write-Host "`tInstalled $($teamsVer)" -ForegroundColor Green 
    } else {
        Write-Host "`tInstallation failed... $($teamsVer)" -ForegroundColor Green 
    }
}

# Update-check 
function Start-TeamsUpdate {

    [CmdletBinding()]
    param (
        [Parameter()][datetime]$timeNow = [datetime]::Now,
        [Parameter()][string]$teamsUpd = "$($env:LOCALAPPDATA)\Microsoft\Teams\update.exe"
    )

    $timeLAT = (Get-Item $teamsUpd).LastAccessTime
    $timeSpan = (New-TimeSpan -Start $timeLAT -End $timeNow).Hours

    Write-Host "`n[info] Update:`n`tCheck for updates if (update.exe).accesstime < 72 hours`n`t Last access: $($timeSpan) hours ago"

    if ($timeSpan -ge 72) {
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
        [Parameter()][string]$teamsExe = "$($env:TEMP)\XXXX\Teams\Teams_windows_x64.exe" # .EXE saved here
    )

    # Write-Host "`n[info] Action:`n`tStopping Teams (if running)..."
    # Get-Process teams -ErrorAction SilentlyContinue | Stop-Process -ErrorAction SilentlyContinue -PassThru

    if (!(Test-Path $teamsExe)) {
        Write-Host "`tDownloading..."
        Invoke-WebRequest -Uri $teamsUrl -OutFile $teamsExe -UseBasicParsing
    } 

    Write-Host "`tInstalling..."

    Start-Process -FilePath $teamsExe -ArgumentList "-s" -Wait -PassThru
    Start-Process -FilePath $teamsUpd -ArgumentList "--processStart teams.exe" -PassThru

    while (!$pcsAge) {
        $teamsPcs = Get-Process teams -ErrorAction SilentlyContinue

        if ($teamsPcs) {
            $pcsAge = $teamsPcs[0].StartTime.AddSeconds(8) -lt [datetime]::Now 
        }
        Start-Sleep 2
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

    if (Test-Path $copySrc -ErrorAction SilentlyContinue) {

        if (!(Test-Path $copyDst)) {
            mkdir $copyDst
        }

        $bgRef = Get-ChildItem $copySrc -Name -ErrorAction SilentlyContinue
        $bgDiff = Get-ChildItem $copyDst -Exclude *thumb* -Name -ErrorAction SilentlyContinue

        foreach ($img in $bgRef) {
            if ($img -notin $bgDiff) {
                (Copy-Item -path $copySrc$img -Destination $copyDst -PassThru).Name  | Out-Host
                $imgCount += 1
            }
        }
        # Update local numbers after job is done
        return Write-Host "`tImages updated: $($imgCount)"
    }
     Write-Host "`tNetwork not available, no images were downloaded." -ForegroundColor Yellow
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

# Activate TeamsAddin for Outlook 
$addinPath = "HKCU:\\SOFTWARE\Microsoft\Office\Outlook\Addins\TeamsAddin.FastConnect"
$addinCheck = (Get-ItemProperty -Path $addinPath -ErrorAction SilentlyContinue).LoadBehavior -ne 3

if ($addinCheck) {
    Write-Host "`n[info] Addin: Enable TeamsAddin for Outlook [ LoadBehavior = 3 ]"
    Set-ItemProperty -Path $addinPath -Name LoadBehavior -Value 3 -ErrorAction SilentlyContinue
}

} # END Measure command

Write-Host "`n[info] Runtime:" 
Write-Host "`t$($timed.TotalSeconds) seconds" -ForegroundColor Cyan

# End logging
Stop-Transcript
