# Variables
$ADPDepartmentField = "Home Department Name"
$DesktopPath        = [Environment]::GetFolderPath("Desktop")
$EmployeeDataFile   = "$DesktopPath\EmployeeData.csv"
$LogPath            = "C:\logs"
$LogFile            = $LogPath + "\" + $(Get-Date -Format FileDate) + "-" + $MyInvocation.MyCommand.Name + ".log"
$OutputPath         = "C:\data"

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

# Check for AD PS Module and set $SearchBase
Try {
    $Domain     = Get-ADDomain -Current LoggedOnUser
    $DeptOU     = "OU=DepartmentGroups,OU=SecurityGroups"
    $SearchBase = "$DeptOU,$Domain" # Set $SearchBase variable
}
Catch {
    Write-Host -ForegroundColor Yellow "Please install RSAT and re-run. (Ref: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools)" 
    Write-Host -ForegroundColor Yellow "Exiting..."

    # Stop Logging
    Stop-Transcript

    Break
}

# Import ADP Data
$AllEmployeeData = Import-Csv $EmployeeDataFile

# Only get 'Active' & 'Leave' employees
$ActiveEmployees = $AllEmployeeData | Where {($_."Status Type" -eq "Active" -or $_."Status Type" -eq "Leave")}

# Sort ADP "Home Department Name" alphabetically and remove duplicates
$UniqueADPDepartments = $ActiveEmployees | Select $ADPDepartmentField | Sort -Property $ADPDepartmentField -Unique

# Collect $ADDepartmentData based on $UniqueADPDepartments
$ADDepartmentData = Foreach ($ADPHomeDepartment in $UniqueADPDepartments.$ADPDepartmentField) {
    
    # Check if $ADPHomeDepartment exists AD
    $HomeDepartmentADGroup = Get-ADGroup -SearchBase $SearchBase -Filter {Name -eq $ADPHomeDepartment} -Properties *
    
    # Set Group Description
    $HomeDepartmentADGroupDescription = "[Maintained by automation] Members of this group are employees in the $ADPHomeDepartment Department according to ADP"

    #If the $ADPHomeDeparmtment exists in AD, write $ADPHomeDepartment to $ADDepartmentData as True
    If ($HomeDepartmentADGroup) {
        Write-Host -ForegroundColor Green "$ADPHomeDepartment exists in $SearchBase."
        [pscustomobject]@{
            GroupName = $ADPHomeDepartment
            Status    = $true
        }

        # If $HomeDepartmentADGroup Description is null, set it
        If (!$($HomeDepartmentADGroup.Description)) {
            Write-Host -ForegroundColor Cyan "$ADPHomeDepartment is missing a Description. Adding it..."
            Set-ADGroup -Identity $ADPHomeDepartment `
                -Description $HomeDepartmentADGroupDescription `
                -ErrorAction Stop
        }
    }

    # If the $ADPHomeDepartment does not exist in AD, write $ADPHomeDepartment to $ADDepartment Data as False
    Elseif (!$HomeDepartmentADGroup) {
        [pscustomobject]@{
            GroupName = $ADPHomeDepartment
            Status    = $false
        }

        Write-Host -ForegroundColor Magenta "$ADPHomeDepartment does not exist. Creating Security Group in $SearchBase..."
        New-ADGroup -Name $ADPHomeDepartment `
            -SamAccountName $ADPHomeDepartment `
            -GroupCategory Security `
            -GroupScope Global `
            -DisplayName $ADPHomeDepartment `
            -Path $SearchBase `
            -Description $HomeDepartmentADGroupDescription `
            -ErrorAction Stop
    }
}

# Export ADP Departments that do not exist in AD to .csv
$ADDepartmentData | `
    Where Status -eq $false | `
    Sort GroupName | `
    Export-Csv -Path C:\data\ADP-Departments-Do-Not-Exist-In-AD.csv -NoTypeInformation

# Collect All $ADDepartmentGroups
$ADDepartmentGroups = Get-ADGroup -SearchBase $SearchBase -Filter *

# Compare $ADDepartmentGroups to $ADPDepartments
$ADDepartmentsMissingFromADP = Foreach ( $Department in $ADDepartmentGroups.Name ) {
    
    # If the $ADDepartmentGroups Name does not exist in $UniqueADPDepartments "Home Department Name" field, write to $DepartmentMissingFromADP as False
    If (!($UniqueADPDepartments.$ADPDepartmentField).Contains($Department)) {
        [pscustomobject]@{
            GroupName = $Department
            Status    = $false
        }
    }
}

# Export AD Departments that do not exist in ADP to .csv
$ADDepartmentsMissingFromADP | `
    Sort GroupName | `
    Export-Csv -Path C:\data\AD-Departments-Do-Not-Exists-In-ADP.csv -NoTypeInformation

# Display session metrics
$MetricsTitle = @'

 #####################
 ## SESSION METRICS ##
 #####################

'@

Write-Host -ForegroundColor Gray $MetricsTitle
Write-Host -ForegroundColor Green "$($($ADDepartmentData.Status -eq $true).Count) ADP Departments exist in $SearchBase"
Write-Host -ForegroundColor Magenta "$($($ADDepartmentData.Status -eq $false).Count) ADP Departments did not exist in $SearchBase"
Write-Host -ForegroundColor Red "$($ADDepartmentsMissingFromADP.Count) AD Department Groups do not exist in the ADP 'Home Department Name' values`n"

# Stop Logging
Stop-Transcript