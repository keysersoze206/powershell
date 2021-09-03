<#
.Synopsis
   Disable User in Active Directory based on Terminated Status Type in ADP
.DESCRIPTION
   Disable User in Active Directory based on Terminated Status Type in ADP
.EXAMPLE
   Disable-TerminatedEmployees.ps1
.EXAMPLE
   Disable-TerminatedEmployees.ps1 -LogFile D:\folderName\output.log `
                                   -EmployeeDataFile C:\folderName\file.csv `
                                   -DateRange LastWeek
#>

Param
(   
    # $LogFile Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.log|^.*.txt")]
    [string]
    $LogFile = $null,

    # $EmployeeDataFile Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.csv")]
    [string]
    $EmployeeDataFile = $null,
    
    # $DateRange Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("LastYear", "LastQuarter", "LastMonth", "LastWeek", "LastDay")] 
    [string]
    $DateRange = $null
)

Begin
{
    Function Timestamp {
        $(Get-Date).TimeOfDay
    }

    Function Write-InfoMsg {
        Write-Host "INFO: [$(Timestamp)] $args"
    }

    Function Write-SuccessMsg {
        # Write message
        Write-Host -ForegroundColor Green "SUCCESS: [$(Timestamp)] $args"
    }

    Function Write-ErrorMsg {
        # Write message
        Write-Host -ForegroundColor Red "ERROR: [$(Timestamp)] $args"
    }

    #Variables
    $Today = $(Get-Date)

    # Reset Counters
    $UserDisabled     = 0
    $UserDoesNotExist = 0
    $UserEnabled      = 0

    # Check for $EmployeeDataFile Param
    Switch ($EmployeeDataFile)
    {
        # Switch: If $EmployeeDataFile was not envoked, use default option
        "$null" 
        {
            $DataFilePath     = [Environment]::GetFolderPath("Desktop")
            $EmployeeDataFile = "EmployeeData.csv"
            $EmployeeDataFile = "$DataFilePath\$EmployeeDataFile"
        }

        # Switch: If $EmployeeDataFile was envoked, use input value
        Default
        {
            Continue
        }
    }

    # Check for $LogFile Param
    Switch ($LogFile)
    {
        # Switch: If $LogFile param was not envoked, use default option
        "$null"
        {
            # Default directory
            $LogPath = "C:\logs"
            # Format Log File as "yyyyMMdd-scriptName.ps1.log"
            $LogFile = "$LogPath\$(Get-Date -Format FileDate)-$($MyInvocation.MyCommand.Name).log"
        }

        # Switch: If $LogFile param was envoked, use input value and set $LogPath
        Default
        {
            $LogPath = Split-Path -Path $LogFile
        }
    }

    # Test if $LogPath directory already exist
    If (!(Test-Path -Path $LogPath)) 
    {
        # Try to create $LogPath
        Try 
        {
            # Create $LogPath
            New-Item -ItemType Directory -Force -Path $LogPath -ErrorAction Stop | Out-Null
        }

        # Unable to create $LogPath directory
        Catch 
        {
            # Error and exit
            Throw "Unable to create directory $LogPath."
        }
    }

    # Start Logging
    Start-Transcript -Path $LogFile
}
Process
{
    # Try to import ADP data
    Try 
    {
        # Import ADP employee data
        $AllEmployees = Import-Csv $EmployeeDataFile
    }

    # Cannot import $EmployeeDataFile
    Catch 
    {
        Write-ErrorMsg "Unable to access $EmployeeDataFile."
        Write-ErrorMsg "Exiting..."

        # Stop Logging
        Stop-Transcript

        Exit 200
    }

    Write-InfoMsg "Starting to run $($MyInvocation.MyCommand.Name)..."

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
        # Switch: Date Range set to 'LastYear'
        'LastYear'
        {
            # Calculate Last Year
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                    -and $_."Status Eff Date" -ge $Today.AddYears(-1) }
        }

        # Switch: Date Range set to 'LastQuarter'
        'LastQuarter'
        {
            # Calculate Last Quarter (3 months)
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                    -and $_."Status Eff Date" -ge $Today.AddMonths(-3) }
        }

        # Switch: Date Range set to 'LastMonth'
        'LastMonth'
        {
            # Calculate Last Month
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                    -and $_."Status Eff Date" -ge $Today.AddMonths(-1) }
        }

        # Switch: Date Range set to 'LastWeek'
        'LastWeek'
        {
            # Calculate Last Week (7 days)
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                    -and $_."Status Eff Date" -ge $Today.AddDays(-7) }
        }

        # Switch: Date Range set to 'LastDay'
        'LastDay'
        {
            # Calculate Last Day
            $TerminatedEmployees = $TerminatedEmployees | `
                Where { $_."Status Eff Date" -le $Today `
                    -and $_."Status Eff Date" -ge $Today.AddDays(-1) }
        }

        # Switch: Date Range set to Default (All Terminated Employees)
        Default
        {
            $TerminatedEmployees = $TerminatedEmployees
        }
    }

    Write-InfoMsg "Termination Date Range is set to $DateRange."
    Write-InfoMsg "Processing $(@($TerminatedEmployees).Count) terminated employee records..."

    # Loop through all termninated employees
    Foreach ( $Employee in $TerminatedEmployees)
    {
        # Join Name fields from ADP data
        $Name = $Employee."First Name" + " " + $Employee."Last Name"

        Try
        {
            # Determine if User exists in AD
            $User = Get-ADUser -Filter { Name -eq $Name }
        }
        Catch
        {
            Write-ErrorMsg "Unabled to access Active Directory."

            # Stop Logging
            Stop-Transcript

            Exit 201
        }

        # User exists and is Enabled
        If($EnabledAccounts = $User | Where Enabled -eq $true) 
        {
            Foreach ($Account in $EnabledAccounts ) 
            {
                Write-InfoMsg "$($Account.distinguishedName) needs to be Disabled."
                $UserEnabled += 1
            }
        }

        # User exists and is Disabled
        Elseif ($DisabledAccounts = $User | Where Enabled -eq $false) 
        {
            Foreach ($Account in $DisabledAccounts ) 
            {
                Write-Verbose "[$(Timestamp)] $($Account.distinguishedName) is already Disabled."
                $UserDisabled += 1
            }
        }

        Else {
            Write-ErrorMsg "$Name does not exist in AD."
            $UserDoesNotExist += 1
        }
    }
}
End
{
    #Formatted Text
    $MetricsTitle = @'

:::::::::::::::::::::
:: SESSION METRICS ::
:::::::::::::::::::::

'@
    $SessionOutput = @"
Processed $(@($TerminatedEmployees).Count) terminated employee records.
$UserDisabled Users are already Disabled.
$UserEnabled Users need to be Disabled.
$UserDoesNotExist Users do not exist in AD.

"@

    # Display session metrics
    Write-Output $MetricsTitle
    Write-Output $SessionOutput

    # Check if -Verbose flag was NOT used
    If (!($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent))
    {
        Write-InfoMsg "You can re-run $($MyInvocation.MyCommand.Name) with -Verbose flag to see full details."
    }

    # Stop Logging
    Stop-Transcript
}