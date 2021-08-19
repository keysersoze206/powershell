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
    [ValidatePattern("((((?:CN|OU)=[^,]+,?)+),)?((DC=[^,]+,?)+)$")]
    [string]
    $SearchBase = $null,

    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.csv")]
    [string]
    $OutputFile = $null
)

Begin
{
    # Help Message
    $HelpMessage = @"
Unable to contact the Domain Controller or the Search Base does not exist.
Check your connectivity or install RSAT and re-run $($MyInvocation.MyCommand.Name).
Reference: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools
Exiting...
"@
    
    # Reset Counters
    $NoEmployeeID        = 0
    $NoSSN               = 0
    $UnverifiedEmployees = 0
    $VerifiedEmployees   = 0

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

    # Check for $LogFile Param. If $null, set it.
    Switch ($LogFile)
    {
        "$null"
        {
            $LogPath = "C:\logs"
            # Format Log File as "yyyyMMdd-scriptName.ps1.log"
            $LogFile = "$LogPath\$(Get-Date -Format FileDate)-$($MyInvocation.MyCommand.Name).log"
        }
        Default
        {
            $LogPath = Split-Path -Path $LogFile
        }
    }

    # Create Log directory
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


    # Start Logging
    Start-Transcript -Path $LogFile

    # Check for $SearchBase Param. If $null, set it. Also test $SearchBase.
    Switch ($SearchBase)
    {
        "$null"
        {
            Try 
            {
                $SearchBase = Get-ADDomain -Current LoggedOnUser
            }

            Catch
            {
                Write-Host -ForegroundColor Red $HelpMessage

                # Stop Logging
                Stop-Transcript

                Exit
            }
        }
        Default
        {
            Try
            {
                Get-ADObject -Identity $SearchBase -ErrorAction Stop | Out-Null
            }
            Catch
            {
                Write-Host -ForegroundColor Red $HelpMessage

                # Stop Logging
                Stop-Transcript

                Exit
            }
        }
    }

    # Check for $OutputFile Param. If $null, set it.
    Switch ($OutputFile)
    {
        "$null"
        {
            $OutputPath = "C:\data"
            # Format Output File as "yyyyMMdd-scriptName.csv"
            $OutputFile = "$OutputPath\$(Get-Date -Format FileDate)-$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).csv"
        }
        Default
        {
            $OutputPath = Split-Path -Path $OutputFile
        }
    }

    # Create Log directory
    If (!(Test-Path -Path $OutputPath)) 
    {
        Try 
        {
            New-Item -ItemType Directory -Force -Path $OutputPath -ErrorAction Stop | Out-Null
        }
    
        Catch 
        {
            Write-Host -ForegroundColor Red "Unable to create directory $OutputPath. Please check $LogFile. Exiting..."
    
            # Stop Logging
            Stop-Transcript
    
            Exit
        }
    }
}
Process
{
    # Import ADP Data
    Try
    {
        $AllEmployeeData = Import-Csv $EmployeeDataFile
    }

    Catch
    {
        Write-Host -ForegroundColor Red "Unable to access $EmployeeDataFile."
        Write-Host -ForegroundColor Red "Exiting..."

        # Stop Logging
        Stop-Transcript
        
        Exit
    }

    # Define "Active" Employees (Status = Active or Leave)
    $ActiveEmployees = $AllEmployeeData | Where {($_."Status Type" -eq "Active" -or $_."Status Type" -eq "Leave")}

    Write-Host -ForegroundColor Gray "Preparing to search for AD Users in $SearchBase..."

    # Loop through $ActiveEmployees and collect $Data
    $Data = Foreach ($Employee in $ActiveEmployees) 
    {
        # Join First and Last Name
        $FullName = $Employee."First Name" + " " + $Employee."Last Name"

        # Get Employee's SSN
        $EmployeeFullSSN = $Employee."Tax ID (SSN)"

        # Parse SSN
        $Split = $EmployeeFullSSN.split("-")
        $Last4EmployeeSSN = [string]$($Split[2..($Split.Length+1)])

        # Employee does not have an SSN in ADP
        If (!$Last4EmployeeSSN) 
        {
            Write-Host -ForegroundColor Magenta "No SSN in ADP for $FullName"

            $NoSSN += 1

            [pscustomobject]@{
                FullName   = $FullName
                EmployeeID = $Last4EmployeeSSN
                Accounts   = "No SSN in ADP"
            }
    
            Continue
        }

        # Try to get user based on Full Name, Last 4 of SSN, and Enabled
        $User = Get-ADUser -Filter {(Name -eq $FullName) -and (EmployeeID -eq $Last4EmployeeSSN) -and (Enabled -eq $true)} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties *

        # If $User is not null
        If ($User) 
        {
            Write-Host -ForegroundColor Green "Found $(@($User).Count) matching account(s) for $FullName with $Last4EmployeeSSN."

            $VerifiedEmployees += 1

            [pscustomobject]@{
                FullName   = $FullName
                EmployeeID = $Last4EmployeeSSN
                Accounts   = "Match with EmployeeID"
            }
        }

        # If $User is null
        Elseif (!$User) 
        {
            Write-Host -ForegroundColor Yellow "Unable to find an account for $FullName matching $Last4EmployeeSSN."

            Write-Host -ForegroundColor Cyan "Trying to find an account for $FullName without an EmployeeID..."

            # Try to get user based on Full Name and Enabled (this is a less accurate search)
            $EnabledUserNoSSN = Get-ADUser -Filter {(Name -eq $FullName) -and (Enabled -eq $true)} `
                -SearchBase $SearchBase `
                -SearchScope Subtree `
                -Properties *

            # Try to get user based on Full Name and Disabled (this is for users who may be disabled due to logon inactivity)
            $DisabledUserNoSSN = Get-ADUser -Filter {(Name -eq $FullName) -and (Enabled -eq $false)} `
                -SearchBase $SearchBase `
                -SearchScope Subtree `
                -Properties *

            # If $EnabledUserNoSSN is not null
            If ($EnabledUserNoSSN) 
            {
                # See if multiple records were returned
                If ($(@($EnabledUserNoSSN).Count) -gt 1) 
                {
                    Write-Host -ForegroundColor Cyan "Found $(@($EnabledUserNoSSN).Count) accounts for $FullName without an EmployeeID."

                    $NoSSNEmployeeMultipleAccounts += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "Multiple Active Accounts, No EmployeeID"
                    }
                }
                # Only one record was returned
                Else
                {
                    Write-Host -ForegroundColor Cyan "Found $(@($EnabledUserNoSSN).Count) account for $FullName without an EmployeeID."

                    $NoSSNEmployees += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "1 Active Account, No EmployeeID"
                    }
                }
            }

            # If $DisabledUserNoSSN is not null
            Elseif ($DisabledUserNoSSN) 
            {
                # See if multiple records were returned
                If ($(@($DisabledUserNoSSN).Count) -gt 1) 
                {
                    Write-Host -ForegroundColor Magenta "Found $(@($DisabledUserNoSSN).Count) Disabled accounts for $FullName without an EmployeeID."

                    $NoSSNEmployeeMultipleDisabledAccounts += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "Multiple Disabled Accounts, No EmployeeID"
                    }
                }
                # Only one record was returned
                Else
                {
                    Write-Host -ForegroundColor Magenta "Found $(@($DisabledUserNoSSN).Count) Disabled account for $FullName without an EmployeeID."

                    $NoSSNEmployeesDisabled += 1

                    [pscustomobject]@{
                        FullName   = $FullName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "1 Disabled Account, No EmployeeID"
                    }
                }
            }

            # If there isn't an employee match for Enabled OR Disabled accounts (i.e. No account)
            Elseif (!$EnabledUserNoSSN -and !$DisabledUserNoSSN) 
            {
                Write-Host -ForegroundColor Red "Unable to find any account(s) for $FullName."
        
                $UnverifiedEmployees += 1

                [pscustomobject]@{
                    FullName   = $FullName
                    EmployeeID = $Last4EmployeeSSN
                    Accounts   = "No Accounts in AD matching Full Name"
                }
            }
        }
    }
}
End
{
    Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
    
    # Write [pscustomobject] in $Data to .csv
    $Data | Export-Csv -Path $OutputFile -NoTypeInformation

    # Formatted Text
    $MetricsTitle = @'

#####################
## SESSION METRICS ##
#####################

'@
    $SessionOutput = @"
There are $(@($ActiveEmployees).Count) employees in ADP.
Found $NoSSN employee(s) without an SSN in ADP.
Found $VerifiedEmployees employee(s) with an EmployeeID attribute and an Active account.
Found $NoSSNEmployees employee(s) without an EmployeeID attribute and an Active account.
Found $NoSSNEmployeeMultipleAccounts employee(s) without an EmployeeID attribute and multiple Active accounts.
Found $NoSSNEmployeesDisabled employee(s) without an EmployeeID attribute and a Disabled account.
Found $NoSSNEmployeeMultipleDisabledAccounts employee(s) without an EmployeeID attribute and multiple Disabled accounts.
Unable to find any accounts for $UnverifiedEmployees employee(s) in AD.
"@
    # Display session metrics
    Write-Host -ForegroundColor Gray $MetricsTitle
    Write-Host -ForegroundColor Gray $SessionOutput

    Stop-Transcript
}