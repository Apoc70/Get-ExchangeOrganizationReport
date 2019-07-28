<# 
    .SYNOPSIS 
    This script fetches Exchange organization configuration data and exports it as Word document.

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 1.0, 2019-07

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 
    http://scripts.granikos.eu

    .DESCRIPTION 
     
    .NOTES 
    Requirements 
    - Windows Server 2012 R2  
    - .NET 4.5
    - Exchange Server Management Shell
    - Word 2013+
    
    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 | Initial community release 

    .PARAMETER SendMail
    Switch to send the zipped archive via email

    .PARAMETER MailFrom
    Sender email address

    .PARAMETER MailTo
    Recipient(s) email address(es)

    .PARAMETER MailServer
    FQDN of SMTP mail server to be used

    .EXAMPLE 
#>
[CmdletBinding()]
param(
  [switch] $SendMail,
  [string] $MailFrom = '',
  [string] $MailTo = '',
  [string] $MailServer = ''
)