# Variables
$LogPath = "C:\logs"
$LogFile = $LogPath + "\" + $(Get-Date -Format FileDate) + "-" + $MyInvocation.MyCommand.Name + ".log"
$Reports = "Report-*.ps1"

# Create Log directory
If (!(Test-Path -Path $LogPath)) {
    Try {
        New-Item -ItemType Directory -Force -Path $LogPath -ErrorAction Stop | Out-Null
    }
    Catch {
        Throw "Unable to create directory $LogPath."
    }
}

# Start Logging
Start-Transcript -Path $LogFile

Foreach ($Report in $(Get-ChildItem -Path $PSScriptRoot -Filter $Reports | Sort Name).Name) {
    Invoke-Expression "$PSScriptRoot\$Report"
}

# Stop Logging
Stop-Transcript