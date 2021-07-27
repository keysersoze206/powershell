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

Write-Host -ForegroundColor Gray "Generating a list of Groups in $($SearchBase.distinguishedName)..."
Try {
    $Groups = Get-ADGroup -Filter * -SearchBase $SearchBase.distinguishedName | `
        Where Name -ne 'Domain Users' | `
        Where Name -ne 'Domain Computers' | `
        Sort Name
}
Catch {
    Write-Host -ForegroundColor Yellow "Unable to generate a list of Groups. Exiting..."

    # Stop Logging
    Stop-Transcript

    Break
}

$Data = Foreach ( $Group in $Groups ) {
    Write-Host -ForegroundColor Gray "Collecting membership of $Group..."
    $Users = Get-ADGroupMember -Identity $Group

    $Users = $Users.Name -join ", "

    If ($Users) {
        [pscustomobject]@{
            GroupName  = $Group.Name
            GroupDN    = $Group.distinguishedName
            MemberName = $Users
        }
    }
}

# Create Output File
Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
$Data | Export-Csv -Path $OutputFile -NoTypeInformation

Write-Host -ForegroundColor Green "The report ran successfully. See output at $OutputFile"

# Stop Logging
Stop-Transcript