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
    $Domain     = Get-ADDomain -Current LoggedOnUser # Set $SearchBase variable
    $InactiveOU = "OU=Inactive Users"
    $InactiveOU = "$InactiveOU,$Domain"
    $SearchBase = $Domain
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

Write-Host -ForegroundColor Gray "Generating a list of Disabled Users in $($SearchBase.distinguishedName)..."
Try {
    $Users = Get-ADUser -Filter {(Enabled -eq 'False')} `
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

Write-Host -ForegroundColor Gray "Determining if Disabled Users are not in $InactiveOU..."
$Data = Foreach ( $User in $Users ){
    $DistinguishedName = $User.DistinguishedName

    $Split = $DistinguishedName.split(",")
    $OU = $Split[1..($Split.Length+1)] -join (",")

    If ($OU -ne $InactiveOU){
        Write-Host -ForegroundColor Gray "$($User.Name) is Disabled and in $OU..."
        [pscustomobject]@{
            Name     = $User.Name
            Username = $User.SAMAccountName
            Enabled  = $User.Enabled
            OU       = $OU
        }
    }
}

Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
$Data | Export-Csv -Path $OutputFile -NoTypeInformation 

Write-Host -ForegroundColor Green "The report ran successfully. See output at $OutputFile"

# Stop Logging
Stop-Transcript