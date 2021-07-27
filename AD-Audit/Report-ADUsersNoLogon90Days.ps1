# Output Variables
$LogPath    = "C:\logs"
$LogFile    = $LogPath + "\" + $(Get-Date -Format FileDate) + "-" + $MyInvocation.MyCommand.Name + ".log"
$OutputPath = "C:\data"
$OutputFile = $OutputPath + "\" + $(Get-Date -Format FileDate) + "-" + [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name) + ".csv"

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

# Create Output directory
If (!(Test-Path -Path $OutputPath)) {
    Try {
        New-Item -ItemType Directory -Force -Path $OutputPath -ErrorAction Stop | Out-Null
    }
    Catch {
        Write-Host -ForegroundColor Red "Unable to create directory $OutputPath. Please check $LogFile. Exiting..."

        # Stop Logging
        Stop-Transcript
    
        Break
    }
}

# Check for AD PS Module
Try {
    $SearchBase = Get-ADDomain -Current LoggedOnUser # Set $SearchBase variable
}
Catch {
    $URL = "(Ref: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools)"
    Write-Host -ForegroundColor Yellow "Please install RSAT and re-run. $URL" 
    Write-Host -ForegroundColor Yellow "Exiting..."

    # Stop Logging
    Stop-Transcript

    Break
}

Write-Host -ForegroundColor Yellow "Starting to run $($MyInvocation.MyCommand.Name)..."

# Set LastLogonDate limit
$90Days = (get-date).adddays(-90)

Write-Host -ForegroundColor Gray "Generating a list of Users who have NOT logged on to $SearchBase since $90Days..."
Try {
    $Users = Get-ADUser `
        -Filter {(LastLogonDate -notlike "*" -or LastLogonDate -le $90days) -and (enabled -eq $true)} `
        -SearchBase $SearchBase.distinguishedName `
        -SearchScope Subtree `
        -Properties * | Sort Name
}
Catch {
    Write-Host -ForegroundColor Yellow "Unable to generate a list of Users. Exiting..."

    # Stop Logging
    Stop-Transcript

    Break
}

$Data = Foreach ( $User in $Users ){
    If (!$User.LastLogonDate) {
        Write-Host -ForegroundColor Gray "$($User.Name) was created on $($User.WhenCreated) and has NEVER logged on..."
    }
    Else {
        Write-Host -ForegroundColor Gray "$($User.Name) has not logged on since $($User.LastLogonDate)..."
    }

    $DistinguishedName = $User.DistinguishedName

    $Split = $DistinguishedName.split(",")
    $OU = $Split[1..($Split.Length+1)] -join (",")

    [pscustomobject]@{
        Name            = $User.Name
        Username        = $User.SAMAccountName
        Enabled         = $User.Enabled
        OU              = $OU
        WhenCreated     = $User.WhenCreated
        LastLogonDate   = $User.LastLogonDate
        PasswordLastSet = $User.PasswordLastSet
    }
}

Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
$Data | Export-Csv -Path $OutputFile -NoTypeInformation

Write-Host -ForegroundColor Green "The report ran successfully. See output at $OutputFile"

# Stop Logging
Stop-Transcript