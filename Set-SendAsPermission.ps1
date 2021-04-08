<# 
    .SYNOPSIS 
    This script adds a single user with send-as permissions to mailboxes which are members of a single security group.

    Version 1.0, 2018-09-01

    Author: Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Please send ideas, comments and suggestions to support@granikos.eu

    .LINK 
    http://scripts.granikos.eu

    .DESCRIPTION 

    This script loops through a membership list of an Active Directory security group.
    A single mailbox is added to each mailbox of the security group members to provide send-as permission.
    
    The script can be used to assign an application account (e.g. CRM, ERP) send-as permission to user mailboxes to send emails AS the user and not as the application.

    .NOTES 
    
    Requirements 
    - Exchange Server 2016 or newer
    - Exchange Online PowerShell connection --> https://go.granikos.eu/ConnectToEXO 
    - Exchange Online PowerShell Module to connect w/ MFA --> https://go.granikos.eu/EXOMFA 
    - Utilizes GlobalFunctions PowerShell Module --> http://bit.ly/GlobalFunctions
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release 

    .PARAMETER SendAsGroup
    This is the name of the Active Directory security group containing all the users where the SendAsUserUpn needs to have send-as permission.

    .PARAMETER SendAsUserUpn
    This is the UserPrincipleName of the user (service account) which will be granted send-ad permission.

    .PARAMETER ExchangeOnline
    Use this switch, if the target mailbox are located in Exchange Online. In this case the script must be executed from within an Exchange Online PowerShell session.

    .EXAMPLE 
    Assign Send-As permission to crmapplication@varunagroup.de for all members of 'CRM-FrontLine' security group. The mailboxes as hosted On-Premises!
    
    .\Set-SendAsPermission.ps1 -SendAsGroup 'CRM-FrontLine' -SendAsUserUpn 'crmapplication@varunagroup.de'

    .EXAMPLE 
    Assign Send-As permission to ax@granikoslabs.eu for all members of 'AX-Sales' security group. All mailboxes are hosted in Exchange Online!
    
    .\Set-SendAsPermission.ps1 -SendAsGroup 'AX-Sales' -SendAsUserUpn 'ax@granikoslabs.eu' -ExchangeOnline
#>  
[cmdletbinding(SupportsShouldProcess)]
Param(
  [string]$SendAsGroup = '', 
  [string]$SendAsUserUpn = '',
  [switch]$ExchangeOnline
)

# Import modules
Import-Module -Name ActiveDirectory

# Import GlobalFunctions
if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
  Import-Module -Name GlobalFunctions
}
else {
  Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
  Write-Warning -Message 'Open an administrative PowerShell session and run Import-Module GlobalFunctions'
  Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
  exit
}

# Some general configuration stuff
$LOG_Information = 0
$LOG_Error = 1
$LOG_Warning = 2
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write('Script started #########################################')

# Some script specific configuration stuff
$GroupCount = 0

# Purge log files depending on LogFileRetention
$logger.Purge()

# MAIN #########################################

# Fetch all group members first 
$SendAsMembers = Get-ADGroupMember -Identity $SendAsGroup -ErrorAction SilentlyContinue

$GroupCount = ($SendAsMembers | Measure-Object).Count

if ($GroupCount -ne 0) {
    # The group is not empty, so let's continue

    $logger.Write(('Group [{0}] contains [{1} members]' -f $SendAsGroup, $GroupCount))

    foreach($User in $SendAsMembers) {
        $Mailbox = $null
        
        $SendAsCount = 0

        $UserUpn = (Get-ADUser -Identity $User.samAccountName).UserPrincipalName

        $logger.Write(('Checking AD user [{0}] - UPN [{1}]' -f $User.samAccountName, ($UserUpn)))

        # Check if user mailbox exists
        $Mailbox = Get-Mailbox -Identity $UserUpn -ErrorAction SilentlyContinue
        
        if($Mailbox -eq $null) {
            $logger.Write(('Mailbox [{0}] does NOT exist' -f $UserUpn),$LOG_Warning)
            continue
        }

        if($ExchangeOnline) {
            # we are using Exchange Online
            
            $SendAsCount = (($Mailbox | Get-RecipientPermission | Where-Object {$_.Trustee -ne 'NT AUTHORITY\SELF' -and $_.Trustee -ne 'NULL SID' -and $_.Trustee -eq $SendAsUserUpn}) | Measure-Object).Count

            if($SendAsCount -eq 0) {
                $logger.Write(('Setting Send-As for [{0}] on mailbox [{1}]' -f $SendAsUserUpn, $UserUpn))
                
                # Configure Send-As permission for $SendasUserUpn on target mailbox $UserUpn in Exchange Online    
                $null = Add-RecipientPermission -Identity $UserUpn -AccessRights SendAs -Trustee $SendAsUserUpn -Confirm:$false
                
            }
            else {
                $logger.Write(('Not action required on mailbox [{0}]' -f $UserUpn)) 
            }
        }
        else {
            # we configure on-premises mailboxes
            $SendAsCount = (($Mailbox | Get-ADPermission | Where-Object {($_.ExtendedRights -like '*send-as*') -and -not ($_.User -like 'NT AUTHORITY\SELF')}) | Measure-Object).Count
            
            if($SendAsCount -eq 0) {
                $logger.Write(('Setting Send-As for [{0}] on mailbox [{1}]' -f $SendAsUserUpn, $UserUpn))

                # Configure Send-As permission for $SendAsUserUpn on target mailbox $UserUpn in an on-premises Exchange organization
                $null = $Mailbox | Add-ADPermission -User $SendAsUserUpn -AccessRights ExtendedRight -ExtendedRights 'Send As'
            }
            else {
                $logger.Write(('Not action required on mailbox [{0}]' -f $UserUpn)) 
            }
        }
    }
  
}
else {
    $logger.Write(('Group [{0}] is emtpy!' -f $SendAsGroup), $LOG_Warning)
}

$logger.Write('Script finished #########################################')

Write-Host 'Script finished. See log file for details.'