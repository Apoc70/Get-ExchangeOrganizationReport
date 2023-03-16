# Get-ExchangeOrganizationReport.ps1

## GitHub

[![license](https://img.shields.io/github/license/Apoc70/Get-ExchangeOrganizationReport.svg)](#)
[![license](https://img.shields.io/github/release/Apoc70/Get-ExchangeOrganizationReport.svg)](#)

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

[![twitter](https://img.shields.io/twitter/follow/stensitzki.svg?label=Twitter%20%40Stensitzki&style=social)](https://twitter.com/stensitzki)
[![linked](https://img.shields.io/badge/LinkedIn-thomasstensitzki-0077B5.svg?logo=LinkedIn)](https://www.linkedin.com/in/thomasstensitzki)
[![blog](https://img.shields.io/badge/blog.granikos.eu-2A6496.svg)](https://blog.granikos.eu)
[![mvpblog](https://img.shields.io/badge/blogs.msmvps.com-2A6496.svg)](https://blogs.msmvps.com/thomastechtalk)

# Follow my Tech Talk video channel or podcast
[![spotify](https://img.shields.io/badge/Spotify-1ED760?&style=?style=plastic&logo=spotify&logoColor=white)](https://open.spotify.com/show/2N49k8CLs0VkkQeGObIlPQ?si=2b6f6c229a9c4f1a)
[![youtube](https://img.shields.io/badge/YouTube-FF0000?style=?style=plastic&logo=youtube&logoColor=white)](https://www.youtube.com/@ThomasStensitzki)
[![techtalk](https://img.shields.io/badge/techtalk.granikos.eu-2A6496.svg)](http://techtalk.granikos.eu)

## Additional Credits

- The script is based on the ADDS_Inventory.ps1 PowerScript by Carl Webster [https://github.com/CarlWebster/ActiveDirectory](https://github.com/CarlWebster/ActiveDirectory)