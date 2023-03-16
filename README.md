
<p align="center">
  <a href="https://twitter.com/stensitzki"><img src="https://img.shields.io/twitter/follow/stensitzki.svg?label=Twitter%20%40Stensitzki&style=social"></a>
  <a href="https://www.linkedin.com/in/thomasstensitzki"><img src="https://img.shields.io/badge/LinkedIn-thomasstensitzki-0077B5.svg?logo=LinkedIn"></a>
  <a href="https://blog.granikos.eu"><img src="https://img.shields.io/badge/blog.granikos.eu-2A6496.svg"></a>
</p>
<p align="center">
<a href="https://www.youtube.com/@ThomasStensitzki"><img src="https://img.shields.io/badge/YouTube-FF0000?style=?style=plastic&logo=youtube&logoColor=white"></a>
<a href="https://open.spotify.com/show/2N49k8CLs0VkkQeGObIlPQ?si=2b6f6c229a9c4f1a"><img src="https://img.shields.io/badge/Spotify-1ED760?&style=?style=plastic&logo=spotify&logoColor=white"></a>
<a href="https://podcasts.apple.com/de/podcast/thomas-tech-community-talk/id1626312145"><img src="https://img.shields.io/badge/apple%20music-F34E68?style=?style=plastic&logo=apple%20music&logoColor=whit"></a>
</p>

# Get-ExchangeOrganizationReport.ps1

This script fetches Exchange organization configuration data and exports it as Word document.

## NOTE

The script is currently under development in version 0.9.

You are welcome to contribute to the PowerShell script development.

## Description

This script reads Exchange Organization data and creates a single Microsoft Word document. A later version will support exporting to an Html file.

The script requires an Exchange Management Shell for Exchange Server 2016 or newer. Older EMS versions are not tested.

A locally installed version of Word is required, as plain Html export is not available.

The default file name is 'Exchange-Org-Report [TIMESTAMP].docx'

Most of the script requires only Exchange admin read-only access for the Exchange organization. Querying address list information requires a membership in the RBAC role _"Address Lists"_.

The script queries hardware information from the Exchange server systems and requires local administrator access to the computer systems.

## Requirements

- Windows Server 2016+, Windows 10
- Exchange Server Management Shell
- Word 2016+
- Required Exchange Role Assignment: _Address Lists_
- PowerShell Script saved with UTF-8 encoding as it contains certain UTF-8 characters

## Revision History

- v0.9 Initial community pre-release
- v0.91 Information about processor cores, memory, and page file size added

## Parameters

### CompanyName

The company name to use on the cover page.

### ExportTo

Target output format for the report.

Valid values: MSWord, Html
Default: MSWord

Html is currently not implemented

### CoverPage

The cover page name for use by Microsoft Word.
Only Word 2010 or newer are supported.
The available cover pages depend on the type of Word setup and locale installed on the system.

The default cover page is Sideline.

### CompanyAddress

Company address to use on the cover page, if the cover page contains an Address field.

### CompanyEMail

Company email address to use on the cover page, if the cover page contains an Email field.

### CompanyFax

Company fax number to use on the cover page, if the cover page contains a Fax field.

### CompanyPhone

Company phone number to use on the cover page, if the cover page contains a Phone field.

### ViewEntireForest

ViewEntireForest switch to set the scope for all Exchange cmdlets to view the entire Exchange Org

### ADForest

Specifies the Active Directory forest object by providing the forest name.

Currently not implemented. Reserved for future use.

### ADDomain

Specifies the Active Directory domain object by providing the doamin name.

Currently not implemented. Reserved for future use.

### IncludedDetails

Switch to include object detail information in the genereted report. Including detailed object information might add a large number of additional pages to the report.

Detailed information is included for the following objects:

- User Role Assignments
- Outlook Web App Policies
- Retention Policy Tags
- Mobile Device Policies
- Address Lists
- Malware Policies
- Transport Rules
- Email Address Policies
- Receive Connectors
- Send Connectors
- Database Availability Groups

### IncludePublicFolders

Switch to include detailed reporting on modern public folder hierarchy.

Using this switch results in an extended run of this scripts depending on the size of public folder hierarchy.

Partially implemented.

### IncludeIntroduction

Switch to include an introductory text at the beginning of the report.

### SendMail

Switch to automatically send the generated report by email.

Currently not implemented.

### MailFrom

Email address of the report sender.

### MailTo

Email address of the report recipient.

### MailServer

Fully qualified domain name (FQDN) of the mail server for sending the report email.

## Examples

``` PowerShell
.\Get-ExchangeOrganizationReport.ps1 -ViewEntireForest:$true
```

Creates a Word report for the local Exchange Organization using the default values defined on the parameters section of the PowerShell script.

``` PowerShell
.\Get-ExchangeOrganizationReport.ps1 -Verbose
```

Creates a Microsoft Word report for the local Exchange Organization with a verbose output to the current PowerShell session.

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

## Additional Credits

- The script is based on the ADDS_Inventory.ps1 PowerScript by Carl Webster [https://github.com/CarlWebster/ActiveDirectory](https://github.com/CarlWebster/ActiveDirectory)