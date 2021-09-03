<#
.Synopsis
   Update Active Directory Data for "Active" Employees in ADP
.DESCRIPTION
   Update Active Directory Data for "Active" Employees in ADP
.EXAMPLE
   Update-ActiveEmployeeData.ps1
.EXAMPLE
   Update-ActiveEmployeeData.ps1 -LogFile D:\folderName\output.log `
                                 -EmployeeDataFile C:\folderName\file.csv `
                                 -SearchBase "OU=name,DC=domain,DC=local" `
                                 -OutputFile D:\folderName\output.csv
#>

[CmdletBinding()]
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

    # $SearchBase Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("((((?:CN|OU)=[^,]+,?)+),)?((DC=[^,]+,?)+)$")]
    [string]
    $SearchBase = $null,

    # $OutputFile Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.csv")]
    [string]
    $OutputFile = $null
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

    # Reset Counters
    $NoSSN                       = 0
    $NoEmployeeNumberUsers       = 0
    $ConfirmedAccounts           = 0
    $ConfirmedChanged            = 0
    $ConfirmedUnChanged          = 0
    $ConfirmedFailed             = 0
    $VerifiableEnabledUsers      = 0
    $VerifiableEnabledChanged    = 0
    $VerifiableEnabledUnChanged  = 0
    $VerifiableDisabledUsers     = 0
    $VerifiableDisabledChanged   = 0
    $VerifiableDisabledUnChanged = 0
    $NoSSNEmployees              = 0
    $NoSSNEmployeesDisabled      = 0
    $UnverifiedEmployees         = 0

    # Help Message (formatted)
    $HelpMessage = @"
Unable to contact the Domain Controller or the Search Base does not exist.
Check your connectivity or install RSAT and re-run $($MyInvocation.MyCommand.Name).
Reference: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools
"@

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

    # Check for $LogFile Param.
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

    # Check for $SearchBase Param.
    Switch ($SearchBase)
    {
        # Switch: If $SearchBase param was not envoked, use default option
        "$null"
        {
            # Try to set $SearchBase
            Try
            {
                # Get logged on User's domain
                $SearchBase = Get-ADDomain -Current LoggedOnUser
            }

            # Unable to set $SearchBase, exit.
            Catch
            {
                Write-InfoMsg $HelpMessage
                Write-InfoMsg "Exiting..."

                # Stop Logging
                Stop-Transcript

                Exit 100
            }
        }

        # Switch: If $SearchBase was envoked, use input value
        Default
        {
            # Test input value for $SearchBase Param
            Try
            {
                # Test if $SearchBase is valid
                Get-ADObject -Identity $SearchBase -ErrorAction Stop | Out-Null
            }

            # $SearchBase test is not valid
            Catch
            {
                Write-InfoMsg $HelpMessage
                Write-InfoMsg "Exiting..."

                # Stop Logging
                Stop-Transcript

                Exit 101
            }
        }
    }

    # Check for $OutputFile Param.
    Switch ($OutputFile)
    {
        # Switch: If $OutputFile was not envoked, use default option
        "$null"
        {
            $OutputPath = "C:\data"
            # Format Output File as "yyyyMMdd-scriptName.csv"
            $ScriptName = "$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name))"
            $OutputFile = "$OutputPath\$(Get-Date -Format FileDate)-$ScriptName.csv"
        }

        # Switch: If $OutputFile was envoked, use input value and separate $OutputPath from $OutputFile
        Default
        {
            $OutputPath = Split-Path -Path $OutputFile
        }
    }

    # If $Output Path does not exist
    If (!(Test-Path -Path $OutputPath))
    {
        # Try to create $OutputPath
        Try
        {
            New-Item -ItemType Directory -Force -Path $OutputPath -ErrorAction Stop | Out-Null
        }

        # If unable to create $OutputPath, exit.
        Catch
        {
            Write-InfoMsg "Unable to create directory $OutputPath. Please check $LogFile."
            Write-InfoMsg "Exiting..."

            # Stop Logging
            Stop-Transcript

            Exit 102
        }
    }
}
Process
{
    # Try to import ADP data
    Try
    {
        $AllEmployeeData = Import-Csv $EmployeeDataFile
    }

    # Cannot import $EmployeeDataFile
    Catch
    {
        Write-InfoMsg "Unable to access $EmployeeDataFile."
        Write-InfoMsg "Exiting..."

        # Stop Logging
        Stop-Transcript

        Exit 200
    }

    # Define "Active" Employees (Status = Active or Leave)
    $ActiveEmployees = $AllEmployeeData | Where-Object {($_."Status Type" -eq "Active" -or $_."Status Type" -eq "Leave")}

    Write-InfoMsg "Searching for $(@($ActiveEmployees).Count) ADP Employees in $SearchBase..."

    # Loop through $ActiveEmployees, collect $Data, and update HR information
    $Data = Foreach ($ActiveEmployee in $ActiveEmployees)
    {
        # Join First and Last Name
        $FullName = $ActiveEmployee."First Name" + " " + $ActiveEmployee."Last Name"

        # Set Employee's SSN
        $EmployeeFullSSN = $ActiveEmployee."Tax ID (SSN)"

        # Parse SSN
        $Split = $EmployeeFullSSN.split("-")
        $Last4EmployeeSSN = [string]$($Split[2..($Split.Length+1)])

        # Set Employee Number
        $EmployeeNumber = $($ActiveEmployee."File Number").TrimStart('0')

        # Employee does not have an SSN in ADP
        If (!$Last4EmployeeSSN)
        {
            Write-ErrorMsg "No SSN in ADP for $FullName!"

            # Add 1 to $NoSSN counter
            $NoSSN += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $null
                Status     = "No SSN in ADP"
            }

            # Move on to the next $ActiveEmployee
            Continue
        }

        # Employee does not have a File Number in ADP
        If (!$EmployeeNumber)
        {
            Write-ErrorMsg "$FullName does not have an EmployeeNumber in ADP!"

            # Add 1 to $NoEmployeeNumberUsers counter
            $NoEmployeeNumberUsers += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $Last4EmployeeSSN
                Status     = "No Employee Number in ADP"
            }

            # Move on to the next $ActiveEmployee
            Continue
        }

        # Get accounts that have a matching Employee Number
        $EmployeeNumberUser = Get-ADUser -Filter {EmployeeNumber -eq $EmployeeNumber} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties *

        # Update the HR information for a User with a matching Employee Number
        $EmployeeNumberUser | ForEach-Object `
        {
            Write-Verbose "[$(Timestamp)] $($_.DistinguishedName) is linked to Employee Number $EmployeeNumber."

            # Add 1 to $ConfirmedAccounts counter
            $ConfirmedAccounts += 1

            [pscustomobject]@{
                FullName   = $_.Name
                DN         = $_.DistinguishedName
                EmployeeID = $Last4EmployeeSSN
                Status     = "Matched by Employee Number"
            }

            $Params = @{ }

            #Determine which User attributes need to be updated; add to Params Hashtable
            If ( $_.EmployeeID -ne $Last4EmployeeSSN )
            {
                $Params.EmployeeID += $Last4EmployeeSSN
            }
            If ( $_.Description -ne $ActiveEmployee."Job Title" )
            {
                $Params.Description += $ActiveEmployee."Job Title"
            }
            If ( $_.Title -ne $ActiveEmployee."Job Title" )
            {
                $Params.Title += $ActiveEmployee."Job Title"
            }
            If ( $_.Department -ne $ActiveEmployee."Home Department Name" )
            {
                $Params.Department += $ActiveEmployee."Home Department Name"
            }

            # If User attributes need to be updated
            If ($Params.Count -gt 0)
            {
                $Changes = $(($Params.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ')
                Try
                {
                    Set-ADUser $_ @Params

                    Write-SuccessMsg "$Changes was modified for $($_.DistinguishedName)."

                    $ConfirmedChanged += 1
                }
                Catch
                {
                    Write-ErrorMsg "Unable to modify $($_.DistinguishedName)."

                    $ConfirmedFailed += 1
                }
            }
            # User attributes are up to date
            Else
            {
                Write-Verbose "[$(Timestamp)] $($_.DistinguishedName) is up to date."

                $ConfirmedUnChanged += 1
            }
        }

        # Get employee accounts that match Full Name and Employee Number is not equal to $EmployeeNumber
        $NoEmployeeNumber = Get-ADUser -Filter {Name -eq $FullName} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties * | `
            Where-Object EmployeeNumber -ne $EmployeeNumber

        # Loop through accounts that do not have an Employee Number, but matches Full Name
        Foreach ($Employee in $NoEmployeeNumber)
        {
            # Check if Employee has a matching "Last 4 of SSN" and is Enabled/Disabled
            Switch ($Employee)
            {
                # Switch: Employee has a matching SSN and is Enabled
                $($Employee | Where-Object {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $true)})
                {
                    Write-Verbose "[$(Timestamp)] Found an Enabled account matching $FullName with EmployeeID $Last4EmployeeSSN."

                    # Set Employee Number
                    $Employee | Where-Object {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $true)} | `
                        ForEach-Object `
                        {
                            Try
                            {
                                Set-ADUser $_ -EmployeeNumber $EmployeeNumber
                                
                                Write-SuccessMsg "Set Employee Number for $($_.DistinguishedName) to $EmployeeNumber."

                                $VerifiableEnabledChanged += 1
                            }
                            Catch
                            {
                                Write-ErrorMsg "Unable to set Employee Number for $($_.DistinguishedName)."

                                $VerifiableEnabledUnChanged += 1
                            }
                        }

                    # Add 1 to VerifiableEnabledUser counter
                    $VerifiableEnabledUsers += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Status     = "Enabled User Match with EmployeeID"
                    }
                }

                # Switch: Employee has a matching SSN and is Disabled
                $($Employee | Where-Object {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $false)})
                {

                    Write-Verbose "[$(Timestamp)] Found a Disabled account matching $FullName with EmployeeID $Last4EmployeeSSN."

                    # Set Employee Number
                    $Employee | Where-Object {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $false)} | `
                        ForEach-Object `
                        {
                            Try
                            {
                                Set-ADUser $_ -EmployeeNumber $EmployeeNumber
                                
                                Write-SuccessMsg "Set Employee Number for $($_.DistinguishedName) to $EmployeeNumber."
                            }
                            Catch
                            {
                                Write-ErrorMsg "Unable to set Employee Number for $($_.DistinguishedName)."
                            }
                        }

                    # Add 1 to VerifiableDisabledUsers counter
                    $VerifiableDisabledUsers += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Status     = "Disabled User Match with EmployeeID"
                    }
                }

                # Switch: Employee does not have a matching SSN and is Enabled
                $($Employee | Where-Object {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $true)})
                {
                    # Set $User
                    # $User = $Employee | Where-Object {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $true)}

                    Write-Verbose "[$(Timestamp)] Found an Enabled account for $FullName but the EmployeeID does not match."

                    # Add 1 to $NoSSNEmployees counter
                    $NoSSNEmployees += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Status     = "Enabled Account, No EmployeeID"
                    }
                }

                # Switch: Employee does not have a matching SSN and is Disabled
                $($Employee | Where-Object {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $false)})
                {
                    # Set $User
                    # $User = $Employee | Where-Object {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $false)}

                    Write-Verbose "[$(Timestamp)] Found a Disabled account for $FullName but the EmployeeID does not match."

                    # Add 1 to NoSSNEmployeesDisabled counter
                    $NoSSNEmployeesDisabled += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Status     = "Disabled Account, No EmployeeID"
                    }
                }
            }
        }

        # Get employee accounts that match Full Name
        $NoEmployeeNumberNoFullName = Get-ADUser -Filter {Name -eq $FullName} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties *

        # If there are not any accounts that match Full Name, move on to the next $ActiveEmployee
        If (!$NoEmployeeNumberNoFullName)
        {
            Write-Verbose "[$(Timestamp)] Unable to find an account for $FullName."

            # Add 1 to $UnverifiedEmployees counter
            $UnverifiedEmployees += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $Last4EmployeeSSN
                Status     = "No Accounts in AD matching Full Name"
            }

            # Move on to the next $ActiveEmployee
            Continue
        }
    }
}
End
{
    Write-InfoMsg "Creating $OutputFile..."

    # Write [pscustomobject] in $Data to .csv
    $Data | Export-Csv -Path $OutputFile -NoTypeInformation

    # Formatted Text
    $MetricsTitle = @'

:::::::::::::::::::::
:: SESSION METRICS ::
:::::::::::::::::::::

'@
    # Formatted Text
    $SessionOutput = @"
There are $(@($ActiveEmployees).Count) Active employees in ADP.
Found $NoSSN employee(s) without an SSN in ADP.
Found $NoEmployeeNumberUsers employee(s) without an Employee Number in ADP.
Found $ConfirmedAccounts confirmed Active Directory Accounts based on the EmployeeNumber attribute.
      $ConfirmedUnChanged Confirmed accounts are up to date.
      $ConfirmedChanged Confirmed accounts have been updated.
      $ConfirmedFailed Confirmed accounts failed to be updated.
Found $VerifiableEnabledUsers employee(s) with an EmployeeID attribute and an Enabled account.
      $VerifiableEnabledChanged Enabled account(s) had EmployeeID added.
      $VerifiableEnabledUnChanged Enabled account(s) failed to add EmployeeID.
Found $VerifiableDisabledUsers employee(s) with an EmployeeID attribute and a Disabled account.
      $VerifiableDisabledChanged Disabled account(s) had EmployeeID added.
      $VerifiableDisabledUnChanged Disabled account(s) failed to add EmployeeID.
Found $NoSSNEmployees employee(s) without an EmployeeID attribute and an Enabled account.
Found $NoSSNEmployeesDisabled employee(s) without an EmployeeID attribute and a Disabled account.
Unable to find any accounts for $UnverifiedEmployees employee(s) in AD.

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

    Exit 0
}