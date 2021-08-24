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
    # Help Message
    $HelpMessage = @"
Unable to contact the Domain Controller or the Search Base does not exist.
Check your connectivity or install RSAT and re-run $($MyInvocation.MyCommand.Name).
Reference: https://docs.microsoft.com/en-us/troubleshoot/windows-server/system-management-components/remote-server-administration-tools
Exiting...
"@

    # Reset Counters
    $NoSSN                   = 0
    $NoEmployeeNumberUsers   = 0
    $ConfirmedAccounts       = 0
    $VerifiableEnabledUsers  = 0
    $VerifiableDisabledUsers = 0
    $NoSSNEmployees          = 0
    $NoSSNEmployeesDisabled  = 0
    $UnverifiedEmployees     = 0

    # Check for $EmployeeDataFile Param
    Switch ($EmployeeDataFile)
    {
        # If $EmployeeDataFile was envoked, set it.
        "$null" 
        {
            $DataFilePath     = [Environment]::GetFolderPath("Desktop")
            $EmployeeDataFile = "EmployeeData.csv"
            $EmployeeDataFile = "$DataFilePath\$EmployeeDataFile"
        }

        #If $EmployeeDataFile was not envoked, use Param value
        Default
        {
            Continue
        }
    }

    # Check for $LogFile Param.
    Switch ($LogFile)
    {
        # If $LogFile param was not envoked, set it.
        "$null"
        {
            # Default directory
            $LogPath = "C:\logs"
            # Format Log File as "yyyyMMdd-scriptName.ps1.log"
            $LogFile = "$LogPath\$(Get-Date -Format FileDate)-$($MyInvocation.MyCommand.Name).log"
        }

        # If $LogFile param was envoked, set $LogPath
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

    # Check for $SearchBase Param. If $null, set it. Also test $SearchBase.
    Switch ($SearchBase)
    {
        # If $SearchBase param was not envoked, use Param value
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
                Write-Host -ForegroundColor Red $HelpMessage

                # Stop Logging
                Stop-Transcript

                Exit
            }
        }

        # If $SearchBase was envoked, set it.
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
        # If $OutputFile is null, use default option
        "$null"
        {
            $OutputPath = "C:\data"
            # Format Output File as "yyyyMMdd-scriptName.csv"
            $OutputFile = "$OutputPath\$(Get-Date -Format FileDate)-$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).csv"
        }

        # Separate Output Path from Output File
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
            Write-Host -ForegroundColor Red "Unable to create directory $OutputPath. Please check $LogFile. Exiting..."
    
            # Stop Logging
            Stop-Transcript
    
            Exit
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
        Write-Host -ForegroundColor Red "Unable to access $EmployeeDataFile."
        Write-Host -ForegroundColor Red "Exiting..."

        # Stop Logging
        Stop-Transcript
        
        Exit
    }

    # Define "Active" Employees (Status = Active or Leave)
    $ActiveEmployees = $AllEmployeeData | Where {($_."Status Type" -eq "Active" -or $_."Status Type" -eq "Leave")}

    Write-Host -ForegroundColor Gray "Preparing to search for AD Users in $SearchBase..."

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
            Write-Host -ForegroundColor DarkRed -BackgroundColor White "No SSN in ADP for $FullName!"

            # Add 1 to $NoSSN counter
            $NoSSN += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $null
                Accounts   = "No SSN in ADP"
            }
            
            # Move on to the next $ActiveEmployee
            Continue
        }

        # Employee does not have a File Number in ADP
        If (!$EmployeeNumber)
        {
            Write-Host -ForegroundColor DarkRed -BackgroundColor White "$FullName does not have an EmployeeNumber in ADP!"
            
            # Add 1 to $NoEmployeeNumberUsers counter
            $NoEmployeeNumberUsers += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $Last4EmployeeSSN
                Accounts   = "No Employee Number in ADP"
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
        $EmployeeNumberUser | % `
        {
            Write-Host "$($EmployeeNumberUser.DistinguishedName) is linked to Employee Number $EmployeeNumber."
            
            # Add 1 to $ConfirmedAccounts counter
            $ConfirmedAccounts += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $_.DistinguishedName
                EmployeeID = $Last4EmployeeSSN
                Status     = "Matched by Employee Number"
            }
        }

        # Get employee accounts that match Full Name and Employee Number is not equal to $EmployeeNumber
        $NoEmployeeNumber = Get-ADUser -Filter {Name -eq $FullName} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties * | `
            Where EmployeeNumber -ne $EmployeeNumber
        
        # If there are not any accounts that match Full Name, move on to the next $ActiveEmployee
        If (!$NoEmployeeNumber)
        {
            Write-Host -ForegroundColor Red "Unable to find a account for $FullName."
        
            # Add 1 to $UnverifiedEmployees counter
            $UnverifiedEmployees += 1

            [pscustomobject]@{
                FullName   = $FullName
                DN         = $null
                EmployeeID = $Last4EmployeeSSN
                Accounts   = "No Accounts in AD matching Full Name"
            }

            # Move on to the next $ActiveEmployee
            Continue
        }

        # Loop through accounts that do not have an Employee Number, but matches Full Name
        Foreach ($Employee in $NoEmployeeNumber)
        {
            # Check if Employee has a matching "Last 4 of SSN" and is Enabled/Disabled
            Switch ($Employee)
            {
                # Switch: Employee has a matching SSN and is Enabled
                $($Employee | Where {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $true)})
                {
                    $User = $Employee | Where {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $true)}
    
                    Write-Host -ForegroundColor Green "Found an Enabled account matching $FullName with $Last4EmployeeSSN."

                    # Set-ADUser $_ -EmployeeNumber $EmployeeNumber -WhatIf | Out-Null
                    
                    # Write-Host -ForegroundColor DarkGreen -BackgroundColor White "Set Employee Number for $($_.DistinguishedName) to $EmployeeNumber."

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
                $($Employee | Where {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $false)})
                {
                    # Set $User
                    $User = $Employee | Where {($_.EmployeeID -eq $Last4EmployeeSSN) -and ($_.Enabled -eq $false)}
    
                    Write-Host -ForegroundColor Yellow "Found a Disabled account matching $FullName with $Last4EmployeeSSN."

                    # Set-ADUser $_ -EmployeeNumber $EmployeeNumber -WhatIf | Out-Null

                    # Write-Host -ForegroundColor DarkYellow -BackgroundColor White "Set Employee Number for $($_.DistinguishedName) to $EmployeeNumber."

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
                $($Employee | Where {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $true)})
                {
                    # Set $User
                    $User = $Employee | Where Enabled -eq $true
    
                    Write-Host -ForegroundColor Cyan "Found an Enabled account for $FullName but the EmployeeID does not match."
    
                    # Add 1 to $NoSSNEmployees counter
                    $NoSSNEmployees += 1
    
                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "1 Enabled Account, No EmployeeID"
                    }
                }
                
                # Switch: Employee does not have a matching SSN and is Disabled
                $($Employee | Where {($_.EmployeeID -ne $Last4EmployeeSSN) -and ($_.Enabled -eq $false)})
                {
                    # Set $User
                    $User = $Employee | Where Enabled -eq $false
    
                    Write-Host -ForegroundColor Magenta "Found a Disabled account for $FullName but the EmployeeID does not match."
                    
                    # Add 1 to NoSSNEmployeesDisabled counter
                    $NoSSNEmployeesDisabled += 1
    
                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = $_.DistinguishedName
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "1 Disabled Account, No EmployeeID"
                    }
                }

                # Switch: Everything else
                Default
                {
                    Write-Host -ForegroundColor Red "Unable to find a account for $FullName. (Default)"
            
                    # Add 1 to $UnverifiedEmployees counter
                    $UnverifiedEmployees += 1
    
                    [pscustomobject]@{
                        FullName   = $FullName
                        DN         = "None"
                        EmployeeID = $Last4EmployeeSSN
                        Accounts   = "No Accounts in AD matching Full Name"
                    }
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
    # Formatted Text
    $SessionOutput = @"
There are $(@($ActiveEmployees).Count) employees in ADP.
Found $NoSSN employee(s) without an SSN in ADP.
Found $NoEmployeeNumberUsers employee(s) with an Employee Number in ADP.
Found $ConfirmedAccounts confirmed Active Accounts based on the EmployeeNumber attribute.
Found $VerifiableEnabledUsers employee(s) with an EmployeeID attribute and an Enabled account.
Found $VerifiableDisabledUsers employee(s) with an EmployeeID attribute and a Disabled account.
Found $NoSSNEmployees employee(s) without an EmployeeID attribute and an Enabled account.
Found $NoSSNEmployeesDisabled employee(s) without an EmployeeID attribute and a Disabled account.
Unable to find any accounts for $UnverifiedEmployees employee(s) in AD.
"@

    # Display session metrics
    Write-Host -ForegroundColor Gray $MetricsTitle
    Write-Host -ForegroundColor Gray $SessionOutput

    # Stop Logging
    Stop-Transcript
}