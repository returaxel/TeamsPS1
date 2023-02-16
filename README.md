Script to manage Teams. Now completely without error handling! 😎

## teamsps1 Features

* Install Teams
* Updates Teams if not done recently
* Download company background images
  * Netshare, locally or from URL
* Activate Outlook addin

*Run in user context*

## teamsfw Features

* Add firewall rules for *main* device user (not ideal for shared devices)

*Run as system/admin*

## Running script
```
Scheduled Task
Packaged as win32 application bundled with backgrounds, for a one-time operation.
Proactive Remediation
```
## Function for fetching background images within a .zip from an URL

Replace the other one with this if needed.

```
### ---------- Background image 
function Get-TeamsPS1Images {
    [CmdletBinding()]
    param (
        [Parameter()][string]$imgUri = "<URL TO .ZIP>",
        [Parameter()][string]$imgDst = "$($env:APPDATA)\Microsoft\Teams\Backgrounds\Uploads",
        [Parameter()][string]$imgZip = "$($env:TEMP)\TeamsPS1\TeamsBG.zip",
        [Parameter()][string]$imgSrc = "$($env:TEMP)\TeamsPS1\TeamsBG\",
        [Parameter()][int]$imgCount = 0
    )
    
    Write-Host "`n[info] Images:"
    # Download ZIP if not found locally
    if (!(Test-Path $imgZip)) {
        Write-Host "`tDownloading images from $($imgUri)"
        try {
            Invoke-WebRequest -Uri $imgUri -OutFile $imgZip -UseBasicParsing
        }
        catch {
            Write-Host "...failed" -ForegroundColor Yellow
        }
    }
    # Extract archive
    if (Test-Path $imgZip) {
        Write-Host "`tExtracting images from .zip archive"
        try {
            Expand-Archive $imgZip -DestinationPath $imgSrc -Force
        }
        catch {
            Write-Host "...failed" -ForegroundColor Yellow
        }
    }
    # Only copy missing background images
    if (Test-Path $imgSrc) {
        Write-Host "`tCopying images"
        if (!(Test-Path $imgDst)) {
            mkdir $imgDst
        }
        $bgRef = Get-ChildItem $imgSrc # -ErrorAction SilentlyContinue
        $bgDiff = Get-ChildItem $imgDst -Exclude *thumb* #-ErrorAction SilentlyContinue
    
        foreach ($img in $bgRef) {
            if ($img.name -notin $bgDiff.Name) {
                (Copy-Item -path $img.FullName -Destination $imgDst -PassThru).Name | Out-Host
                $imgCount += 1
            }
        }
    } else {
        return Write-Host "`tImages updated: ...failed"
    }
    # Update local numbers after job is done
    return Write-Host "`tImages updated: $($imgCount)"
}
```

### Disclaimer

$ErrorActionPreference = Run at your own risk 🤞
