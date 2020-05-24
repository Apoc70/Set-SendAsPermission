# Set-SendAsPermission.ps1

This script adds a single user with send-as permissons to mailboxes which are members of a single security group.

## Description

This script loops through a membership list of an Active Directory security group. A single mailbox is added to each mailbox of the security group members to provide send-as permission.

The script can be used to assign an application account (e.g. CRM, ERP) send-as permission to user mailboxes to send emails AS the user and not as the application.

## Requirements

- Exchange Server 2016 or newer
- Exchange Online PowerShell connection --> [https://go.granikos.eu/ConnectToEXO](https://go.granikos.eu/ConnectToEXO)
- Exchange Online PowerShell Module to connect w/ MFA --> [https://go.granikos.eu/EXOMFA](https://go.granikos.eu/EXOMFA)
- Utilizes GlobalFunctions PowerShell module --> [http://bit.ly/GlobalFunctions](http://bit.ly/GlobalFunctions)

## Parameters

### SendAsGroup

This is the name of the Active Directory security group containing all the users where the SendAsUserUpn needs to have send-as permission.

### SendAsUserUpn

This is the UserPrincipleName of the user (service account) which will be granted send-ad permission.

### ExchangeOnline

Use this switch, if the target mailbox are located in Exchange Online. In this case the script must be executed from within an Exchange Online PowerShell session.

## Examples

``` PowerShell
.\Set-SendAsPermission.ps1 -SendAsGroup 'CRM-FrontLine' -SendAsUserUpn 'crmapplication@varunagroup.de'
```

Assign Send-As permission to crmapplication@varunagroup.de for all members of 'CRM-FrontLine' security group. The mailboxes as hosted On-Premises!

``` PowerShell
.\Set-SendAsPermission.ps1 -SendAsGroup 'AX-Sales' -SendAsUserUpn 'ax@granikoslabs.eu' -ExchangeOnline
```

Assign Send-As permission to ax@granikoslabs.eu for all members of 'AX-Sales' security group. All mailboxes are hosted in Exchange Online!

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)