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

# Get OUs in $SearchBase
Write-Host -ForegroundColor Gray "Generating a list of OUs in $($SearchBase.distinguishedName)..."
Try{
    $OUs = Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase.distinguishedName -SearchScope Subtree
}
Catch {
    Write-Host -ForegroundColor Yellow "Unable to generate a list of OUs. Exiting..."

    # Stop Logging
    Stop-Transcript

    Break
}

# Loop through OUs
$Data = Foreach ($OU in $OUs) {
    
    # Get GUID for GPOs Linked to OU
    Write-Host -ForegroundColor Gray "Collecting info on GPOs Linked to $OU..."
    $GPOGUIDs = $OU.LinkedGroupPolicyObjects | foreach {
        $_.Substring(4,36)
    }

    # Get GPO information
    $GPOName = ForEach ($GPOGUID in $GPOGUIDs) {
        Try {
            Get-GPO -Guid $GPOGUID
        }
        Catch {
            Write-Host -ForegroundColor Yellow "Unable to get GPO for GUID $GPOGUID"
        }
    }
    
    # Join GPO Display Names if multiple GPOs are assigned to an OU
    $GPOName = $GPOName.DisplayName -join ", "

    # Set GPOName to "none" if null
    If (!$GPOName) { $GPOName = "none" }

    # Add GPO info to PSObject
    [pscustomobject]@{
        OU      = $OU.distinguishedName
        GPOName = $GPOName
    }
}

# Create Output File
Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
$Data | Export-Csv -Path $OutputFile -NoTypeInformation

Write-Host -ForegroundColor Green "The report ran successfully. See output at $OutputFile"

# Stop Logging
Stop-Transcript