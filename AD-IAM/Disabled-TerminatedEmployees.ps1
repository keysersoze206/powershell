<#
.Synopsis
   Disable User in Active Directory based on Terminated Status Type in ADP
.DESCRIPTION
   Long description
.EXAMPLE
   Disable-TerminatedEmployees.ps1
.EXAMPLE
   Disable-TerminatedEmployees.ps1 -LogFile D:\folderName\output.log -EmployeeDataFile C:\folderName\file.csv -DateRange LastWeek
#>

Param
(   
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.log|^.*.txt")]
    [string]
    $LogFile = $null,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.csv")]
    [string]
    $EmployeeDataFile = $null,
    
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("LastYear", "LastMonth", "LastWeek", "LastDay")] 
    [string]
    $DateRange = $null
)

Begin
{
    #Variables
    $Today = $(Get-Date)

    # Reset Counters
    $UserDisabled     = 0
    $UserDoesNotExist = 0
    $UserEnabled      = 0

    # Check for $EmployeeDataFile Param
    Switch ($EmployeeDataFile)
    {
        "$null" 
        {
            $DataFilePath     = [Environment]::GetFolderPath("Desktop")
            $EmployeeDataFile = "EmployeeData.csv"
            $EmployeeDataFile = "$DataFilePath\$EmployeeDataFile"
        }
        Default
        {
            Continue
        }
    }

    # Check for $LogFile Param
    Switch ($LogFile) {
        "$null"
        {
            $LogPath = "C:\logs"
            $LogFile = "$LogPath\$(Get-Date -Format FileDate)-$($MyInvocation.MyCommand.Name).log"
        }
        Default
        {
            $LogPath = Split-Path -Path $LogFile
        }
    }

    # Create Log directory
    If ($LogPath)
    {
        If (!(Test-Path -Path $LogPath)) 
        {
            Try 
            {
                New-Item -ItemType Directory -Force -Path $LogPath -ErrorAction Stop | Out-Null
            }

            Catch 
            {
                Throw "Unable to create directory $LogPath."
            }
        }
    }

    # Start Logging
    Start-Transcript -Path $LogFile

    Try 
    {
        # Import ADP employee data
        $AllEmployees = Import-Csv $EmployeeDataFile
    }

    Catch 
    {
        Write-Host -ForegroundColor Yellow "Unable to access $EmployeeDataFile."
        Write-Host -ForegroundColor Yellow "Exiting..."

        # Stop Logging
        Stop-Transcript

        Break
    }
}
Process
{
    Write-Host -ForegroundColor Yellow "Starting to run $($MyInvocation.MyCommand.Name)..."

    # Convert "Status Eff Date" from [string] to [datetime] format
    Foreach ($Employee in $AllEmployees) 
    {
        $Employee."Status Eff Date" = [datetime]::parseexact($Employee."Status Eff Date", 'MM/dd/yyyy', $null)
    }

    # Generate a list of Employees with "Terminated" Status
    $TerminatedEmployees = $AllEmployees | `
        Where  { $_."Status Type" -eq "Terminated" } | `
        Sort -Descending "Status Eff Date"
    
    # Set Termination Date Contraint
    Switch ($DateRange)
    {
        'LastYear'
        { 
            Write-Host -ForegroundColor Gray "Termination Date Range is set to $DateRange."

            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                -and $_."Status Eff Date" -ge $Today.AddYears(-1) }

            Write-Host -ForegroundColor Gray "Processing $(@($TerminatedEmployees).Count) terminated employee records."
        }
        'LastMonth'
        {
            Write-Host -ForegroundColor Gray "Termination Date Range is set to $DateRange."

            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                -and $_."Status Eff Date" -ge $Today.AddMonths(-1) }

            Write-Host -ForegroundColor Gray "Processing $(@($TerminatedEmployees).Count) terminated employee records..."
        }
        'LastWeek'
        {
            Write-Host -ForegroundColor Gray "Termination Date Range is set to $DateRange."
            
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                -and $_."Status Eff Date" -ge $Today.AddDays(-7) }

            Write-Host -ForegroundColor Gray "Processing $(@($TerminatedEmployees).Count) terminated employee records..."
        }
        'LastDay'
        {
            Write-Host -ForegroundColor Gray "Termination Date Range is set to $DateRange."

            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                -and $_."Status Eff Date" -ge $Today.AddDays(-1) }

            Write-Host -ForegroundColor Gray "Processing $(@($TerminatedEmployees).Count) terminated employee records..."
        }
        Default
        {
            Write-Host -ForegroundColor Gray "Termination Date Range is set to All Records (Default)."

            $TerminatedEmployees = $TerminatedEmployees

            Write-Host -ForegroundColor Gray "Processing $(@($TerminatedEmployees).Count) terminated employee records..."
        }
    }

    # Loop through all termninated employees
    Foreach ( $Employee in $TerminatedEmployees)
    {
        # Join Name fields from ADP data
        $Name = $Employee."First Name" + " " + $Employee."Last Name"

        # Determine if User exists in AD
        $User = Get-ADUser -Filter { Name -eq $Name }

        # User exists and is Enabled
        If($EnabledAccounts = $User | Where Enabled -eq $true) 
        {
            Foreach ($Account in $EnabledAccounts ) 
            {
                Write-Host -ForegroundColor Red "$($Account.distinguishedName) needs to be Disabled."
                $UserEnabled += 1
            }
        }

        # User exists and is Disabled
        Elseif ($DisabledAccounts = $User | Where Enabled -eq $false) 
        {
            Foreach ($Account in $DisabledAccounts ) 
            {
                Write-Host -ForegroundColor Green "$($Account.distinguishedName) is already Disabled."
                $UserDisabled += 1
            }
        }

        Else {
            Write-Host -ForegroundColor Yellow "$Name does not exist in AD."
            $UserDoesNotExist += 1
        }
    }
}
End
{
    #Formatted Text
    $MetricsTitle = @'

#####################
## SESSION METRICS ##
#####################

'@

    # Display session metrics
    Write-Host -ForegroundColor Gray $MetricsTitle
    Write-Host -ForegroundColor Gray "Processed $(@($TerminatedEmployees).Count) terminated employee records."
    Write-Host -ForegroundColor Green "$UserDisabled Users are already Disabled."
    Write-Host -ForegroundColor Red "$UserEnabled Users need to be Disabled."
    Write-Host -ForegroundColor Yellow "$UserDoesNotExist Users do not exist in AD."

    # Stop Logging
    Stop-Transcript
}