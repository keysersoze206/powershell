<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
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
    
            Break
        }
    }
}
Process
{
    $SearchBase = Get-ADDomain -Current LoggedOnUser
    $SearchBase = "OU=Administrator Accounts,$SearchBase"
    
    $AdminUsers = Get-ADUser -Filter * `
        -SearchBase $SearchBase `
        -SearchScope Subtree `
        -Properties Name, DistinguishedName, MemberOf |`
        Sort Name
    
    $Data = Foreach ($Admin in $AdminUsers)
    {
        $AdminGroups = $Admin | Select -ExpandProperty MemberOf

        $AdminGroups = Foreach ($AdminGroup in $AdminGroups)
        {
            Get-ADObject -Identity $AdminGroup
        }
        $AdminGroups = $AdminGroups.Name -join ", "
    
        [pscustomobject]@{
            Name     = $Admin.Name
            Username = $Admin.SAMAccountName
            DN       = $Admin.DistinguishedName
            Enabled  = $Admin.Enabled
            MemberOf = $AdminGroups
        }
    }
}
End
{
    $Data | Export-Csv -Path $OutputFile -NoTypeInformation

    Stop-Transcript
}