<#
.Synopsis
   Generate a list of Users in Active Directory who haven't logged in for a period of time
.DESCRIPTION
   Generate a list of Users in Active Directory who haven't logged in for a period of time
.EXAMPLE
   Get-ADUsersNoRecentLogon.ps1
.EXAMPLE
   Get-ADUsersNoRecentLogon.ps1 -LogFile D:\folderName\output.log `
                                -SearchBase "OU=name,DC=domain,DC=local" `
                                -OutputFile D:\folderName\output.csv
                                -DateRange 90Days
#>

Param
(
    # $LogFile Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern("^.*.log|^.*.txt")]
    [string]
    $LogFile = $null,

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
    $OutputFile = $null,
    
    # $DateRange Param
    [Parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("30Days", "60Days", "90Days", "180Days")] 
    [string]
    $DateRange = $null
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
    
    Write-Host -ForegroundColor Yellow "Starting to run $($MyInvocation.MyCommand.Name)..."

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
                Write-Host -ForegroundColor Red $HelpMessage

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
                Write-Host -ForegroundColor Red $HelpMessage

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
            $OutputFile = "$OutputPath\$(Get-Date -Format FileDate)-$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)).csv"
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
            Write-Host -ForegroundColor Red "Unable to create directory $OutputPath. Please check $LogFile. Exiting..."
    
            # Stop Logging
            Stop-Transcript
    
            Exit 102
        }
    }

    # Set Termination Date Contraint
    Switch ($DateRange)
    {
        # Switch: Date Range set to '30Days'
        '30Days'
        {
            $Range = 30

            Write-Host -ForegroundColor Gray "Date Range is set to '$($Range)Days'..."
        }

        # Switch: Date Range set to '60Days'
        '60Days'
        {
            $Range = 60

            Write-Host -ForegroundColor Gray "Date Range is set to '$($Range)Days'..."
        }

        # Switch: Date Range set to '90Days'
        '90Days'
        {
            $Range = 90

            Write-Host -ForegroundColor Gray "Date Range is set to '$($Range)Days'..."
        }

        # Switch: Date Range set to '180Days'
        '180Days'
        {
            $Range = 180

            Write-Host -ForegroundColor Gray "Date Range is set to '$($Range)Days'..."
        }

        # Switch: Date Range set to Default (90Days)
        Default
        {
            $Range = 90

            Write-Host -ForegroundColor Gray "Using Default Date Range: '$($Range)Days'..."
        }
    }
}
Process
{
    # Set LastLogonDate limit
    $Date = (get-date).adddays(-$Range)
    
    Write-Host -ForegroundColor Gray "Generating a list of Users in $SearchBase who have NOT logged on since $Date..."
    
    Try
    {
        $Users = Get-ADUser `
            -Filter {(LastLogonDate -notlike "*" -or LastLogonDate -le $Date) -and (enabled -eq $true)} `
            -SearchBase $SearchBase `
            -SearchScope Subtree `
            -Properties * | `
            Sort Name
    }
    Catch
    {
        Write-Host -ForegroundColor Yellow "Unable to generate a list of Users. Exiting..."
    
        # Stop Logging
        Stop-Transcript
    
        Exit 200
    }
    
    $Data = Foreach ( $User in $Users )
    {
        If (!$User.LastLogonDate)
        {
            Write-Host -ForegroundColor Gray "$($User.Name) was created on $($User.WhenCreated) and has NEVER logged on..."
        }
        Else
        {
            Write-Host -ForegroundColor Gray "$($User.Name) has not logged on since $($User.LastLogonDate)..."
        }
        
        # Regular Expression for to find CN=Users and any OU
        [regex]$regex="((((?:[CN=Users])|(OU)=[^,]+,?)+),)?((DC=[^,]+,?)+)$"
        # Parse OU from User's DistinguishedName
        $OU = $($regex.Match($User.DistinguishedName)).Value
        
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
}
End
{
    Write-Host -ForegroundColor Yellow "Creating $OutputFile..."
    
    $Data | Export-Csv -Path $OutputFile -NoTypeInformation
    
    Write-Host -ForegroundColor Green "The report ran successfully. See output at $OutputFile"
    
    # Stop Logging
    Stop-Transcript
}