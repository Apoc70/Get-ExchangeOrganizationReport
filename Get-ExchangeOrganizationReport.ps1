<# 
    .SYNOPSIS 

    This script fetches Exchange organization configuration data and exports it as Word document.

    Thomas Stensitzki 

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

    Version 0.9, 2019-10

    Please send ideas, comments and suggestions to support@granikos.eu 

    .LINK 

    http://scripts.granikos.eu

    .DESCRIPTION 

    This script reads Exchange Organization 

    To fetch address list information the accoutn executing this script requires a role
    group that has the 'Address Lists' role assgined.

    The script is based on the ADDS_Inventory-ps1 PowerScript by 
     
    .NOTES 
    
    Requirements 
    - Windows Server 2016+, Windows 10
    - Exchange Server Management Shell
    - Word 2016+
    - Role Assignment: Address Lists

    
    Revision History 
    -------------------------------------------------------------------------------- 
    0.9 | Initial community release 

    .PARAMETER 

    .EXAMPLE 
#>
[CmdletBinding()]
param(
  [string]$CompanyName = 'ACME Inc.',
  [ValidateSet('MSWord','Html')]
  $ExportTo = 'MSWord',
  [string]$CoverPage = 'Sideline',
  [string]$CompanyAddress = 'ACME Street, ACME City, 55555',
  [string]$CompanyEmail = 'email@mcsmemail.de',
  [string]$CompanyFax = '+XX FAX',
  [string]$CompanyPhone = '+XX PHONE',
  [switch]$ViewEntireForest,
  [string]$ADForest = $Env:USERDNSDOMAIN,
  [string]$ADDomain = '',
  [switch]$IncludeDetails,
  [switch]$IncludePublicFolders,
  [switch]$IncludeIntroduction,
  [switch]$SendMail,
  [string]$MailFrom = '',
  [string]$MailTo = '',
  [string]$MailServer = ''
)

# Some variables to declare
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
[Diagnostics.Stopwatch]$StopWatch =  [Diagnostics.Stopwatch]::StartNew()
[string]$FileName = 'Exchange-Org-Report'

# Save current error action preference to restore the setting when script finishes
$SavedErrerActionPreference = $ErrorActionPreference
# Set error action preference
$ErrorActionPreference = 'SilentlyContinue'

# Default values
$NA = 'N/A'
$NotExported = 'Not yetexported to Word'
$GeneratedOn = (Get-Date -f yyyy-MM-dd)
$ReportCulture = "de-DE"

#region Type Definition

try {

  Add-Type -TypeDefinition @"

  using System.Collections;

  namespace OrgReport {
    
    public class ExchangeServerObject {
      public string Name;
      public OperatingSystemObject OperatingSystem;
    }

    public class OperatingSystemObject {
      public OSVersionName  OSVersion; 
      public string OSVersionBuild; 
      public string OSName;
      public object OperatingSystem;
      public string BootUpTimeInDays;
      public string BootUpTimeInHours;
      public string BootUpTimeInMinutes;
      public string BootUpTimeInSeconds;
      public string TimeZone;
      public bool PendingReboot;
      public System.Array TLSSettings;
    }

    //enum for OSVersion
    public enum OSVersionName {
      Unknown,
      Windows2008, 
      Windows2008R2,
      Windows2012,
      Windows2012R2,
      Windows2016,
      Windows2019
    }

    public enum ExchangeCULevel {
      Unknown,
      Preview,
      RTM,
      CU1,
      CU2,
      CU3,
      CU4,
      CU5,
      CU6,
      CU7,
      CU8,
      CU9,
      CU10,
      CU11,
      CU12,
      CU13,
      CU14,
      CU15,
      CU16,
      CU17,
      CU18,
      CU19,
      CU20,
      CU21,
      CU22,
      CU23
    }
  }

"@
}
catch {
  # oops 
  Write-Warning -Message 'The script was unable to add custom classes to this PowerShell session. Please close this PowerShell session and open a new session.'
  exit 
}
finally {
  # restore saved preferred error action
    $ErrorActionPreference = $SavedErrerActionPreference
}

#end region

function Stop-Script {

  if($ExportTo -eq 'MSWord') {
    # Cleanup ComObject
    $Script:Word.Quit()
    Write-Verbose "$(Get-Date): System Cleanup"

    [Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null

    if(Test-Path variable:global:word) {
      Remove-Variable -Name Word -Scope Global -Force -Confirm:$false
    }
  }
  # Call Garbage Collector
  [gc]::Collect() 
  [gc]::WaitForPendingFinalizers()
  Write-Verbose "$(Get-Date): Script has been aborted"
  $ErrorActionPreference = $SavedErrerActionPreference
  Exit
}

function Show-ProgressBar {
  [CmdletBinding()]
  param(
    [int]$PercentComplete,
    [string]$Status = '',
    [int]$Stage,
    [string]$Activity = 'Get-ExchangeOrganizationReport'
  )

  $TotalStages = 5
  Write-Progress -Id 1 -Activity $Activity -Status $Status -PercentComplete (($PercentComplete/$TotalStages)+(1/$TotalStages*$Stage*100))
}

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue {
  param(
    [string]$Path, 
    [string]$Name
  )
  $key = Get-Item -LiteralPath $Path -EA 0
  $key -and $Null -ne $key.GetValue($Name, $Null)
}

# Gets the specified local registry value or $Null if it is missing
Function Get-LocalRegistryValue {
  param (
    [string]$Path, 
    [string]$Name
  )
  $key = Get-Item -LiteralPath $Path -ea 0
  if($key) {
    $key.GetValue($Name, $Null)
  }
  else {
    $Null
  }
}

function Test-CompanyName {
  $RegistryPath = 'HKCU:\Software\Microsoft\Office\Common\UserInfo'
  [bool]$Result = Test-RegistryValue -Path $RegistryPath -Name 'CompanyName'

  if($Result) {
    Return Get-LocalRegistryValue -Path $RegistryPath -Name "CompanyName"
  }
  else {
    $Result = Test-RegistryValue -Path $RegistryPath -Name "Company"
		
    if($Result) {
      Return Get-LocalRegistryValue -Path $RegistryPath -Name "Company"
    }
    else {
      Return ''
    }
  }
}
Function Get-RegistryValue {	
  [CmdletBinding()]
  Param(
    [string]$Path, 
    [string]$Name, 
    [string]$ComputerName
  )
  # Gets the specified registry value or $Null if it is missing

  if($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq "LocalHost")	{
    $key = Get-Item -LiteralPath $path -ea 0
    if($key) {
      Return $key.GetValue($Name, $Null)
    }
    else {
      Return $Null
    }
  }
  else {
    #path needed here is different for remote registry access
    $path1 = $Path.SubString(6)
    $path2 = $Path1.Replace('\','\\')
    $Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
    $RegKey= $Reg.OpenSubKey($path2)
    $Results = $RegKey.GetValue($Name)

    if($Null -ne $Results) {
      Return $Results
    }
    else {
      Return $Null
    }
  }
}

#endregion 

#region Word functions

function Test-WordPrerequisites {

  if(!(Test-Path -Path REGISTRY::HKEY_CLASSES_ROOT\Word.Application)) {

    # Word is not installed
    $ErrorActionPreference = $SavedErrerActionPreference
		Write-Warning -Message 'This script uses Microsoft Word, please install Microsoft Word or run the script on a computer with Microsoft Word installed.'

    # exit script
		Exit
  }
  else {
    # get current session id
    $SessionID = (Get-Process -PID $PID).SessionId
    
    [bool]$IsRunning = ($null -ne ((Get-Process -Name 'WinWord' -ErrorAction SilentlyContinue) | Where-Object {$_.SessionId -eq $SessionID})	)

	  if($IsRunning) {

      # There is an active Word instance
		  $ErrorActionPreference = $SavedErrerActionPreference
		  Write-Warning -Message "Please close all running instances of Microsoft Word and restart the PowerShell script"

      # exit script
		  Exit
	  }
  }
}

function Set-WordHashTable {
  Param([string]$CultureCode)

  #optimized by Michael B. SMith
	
  # DE and FR translations for Word 2010 by Vladimir Radojevic
  # Vladimir.Radojevic@Commerzreal.com

  # DA translations for Word 2010 by Thomas Daugaard
  # Citrix Infrastructure Specialist at edgemo A/S

  # CA translations by Javier Sanchez 
  # CEO & Founder 101 Consulting

  #ca - Catalan
  #da - Danish
  #de - German
  #en - English
  #es - Spanish
  #fi - Finnish
  #fr - French
  #nb - Norwegian
  #nl - Dutch
  #pt - Portuguese
  #sv - Swedish
  #zh - Chinese

  [string]$toc = $(
    Switch ($CultureCode)
    {
      'ca-'	{ 'Taula automática 2'; Break }
      'da-'	{ 'Automatisk tabel 2'; Break }
      'de-'	{ 'Automatische Tabelle 2'; Break }
      'en-'	{ 'Automatic Table 2'; Break }
      'es-'	{ 'Tabla automática 2'; Break }
      'fi-'	{ 'Automaattinen taulukko 2'; Break }
      'fr-'	{ 'Table automatique 2'; Break } #changed 13-feb-2017 david roquier and samuel legrand
      'nb-'	{ 'Automatisk tabell 2'; Break }
      'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
      'pt-'	{ 'Sumário Automático 2'; Break }
      'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
      'zh-'	{ '自动目录 2'; Break }
    }
  )

  $Script:myHash                      = @{}
  $Script:myHash.Word_TableOfContents = $toc
  $Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
  $Script:myHash.Word_Heading1        = $wdStyleheading1
  $Script:myHash.Word_Heading2        = $wdStyleheading2
  $Script:myHash.Word_Heading3        = $wdStyleheading3
  $Script:myHash.Word_Heading4        = $wdStyleheading4
  $Script:myHash.Word_TableGrid       = $wdTableGrid

}

Function Write-WordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
  Param(
    [int]$Style=0, 
    [int]$Tabs = 0, 
    [string]$Name = '', 
    [string]$Value = '', 
    [string]$FontName=$Null,
    [int]$FontSize=0,
    [bool]$Italics=$False,
    [bool]$Boldface=$False,
  [Switch]$NoNewLine)
	
  #Build output style
  [string]$output = ""
  Switch ($style) {
    0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
    1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
    2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
    3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
    4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
    Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
  }
	
  #build # of tabs
  While($tabs -gt 0) 	{ 
    $output += "`t"
    $tabs--
  }
 
  if(![String]::IsNullOrEmpty($fontName)) {
    $Script:Selection.Font.name = $fontName
  } 

  if($fontSize -ne 0) 	{
    $Script:Selection.Font.size = $fontSize
  } 
 
  if($italics -eq $True) {
    $Script:Selection.Font.Italic = $True
  } 
 
  if($boldface -eq $True) {
    $Script:Selection.Font.Bold = $True
  } 

  #output the rest of the parameters.
  $output += $name + $value
  $Script:Selection.TypeText($output)
 
  #test for new WriteWordLine 0.
  if(!($nonewline)) {
    $Script:Selection.TypeParagraph()
  } 
}

function Test-WordCoverPage {
  [CmdletBinding()]
  Param(
    [int]$WordVersion, 
    [string]$CoverPage, 
    [string]$CultureCode
  )
	
  $CoverPageArray = ""
	
  Switch ($CultureCode)	{
    'ca-'	{
      if($WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "En bandes", "Faceta", "Filigrana",
          "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
          "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
        "Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
      }
      elseif($WordVersion -eq $wdWord2013) {
        $CoverPageArray = ("Austin", "En bandes", "Faceta", "Filigrana",
          "Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
          "Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
        "Sector (fosc)", "Semàfor", "Visualització", "Whisp")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alfabet", "Anual", "Austin", "Conservador",
          "Contrast", "Cubicles", "Diplomàtic", "Exposició",
          "Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
          "Perspectiva", "Piles", "Quadrícula", "Sobri",
        "Transcendir", "Trencaclosques")
      }
    }

    'da-'	{
      if($WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Bevægelse", "Brusen", "Facet", "Filigran", 
          "Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
          "Retro", "Semafor", "Sidelinje", "Stribet", 
        "Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
      }
      elseif($WordVersion -eq $wdWord2013) {
        $CoverPageArray = ("Bevægelse", "Brusen", "Ion (lys)", "Filigran",
          "Retro", "Semafor", "Visningsmaster", "Integral",
          "Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
        "Udsnit (mørk)", "Ion (mørk)", "Austin")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Bevægelse", "Moderat", "Perspektiv", "Firkanter",
          "Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
          "Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
        "Nålestribet", "Årlig", "Avispapir", "Tradionel")
      }
    }

    'de-'	{
      if($WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Bewegung", "Facette", "Filigran", 
          "Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
          "Pfiff", "Randlinie", "Raster", "Rückblick", 
          "Segment (dunkel)", "Segment (hell)", "Semaphor", 
        "ViewMaster")
      }
      elseif($WordVersion -eq $wdWord2013) {
        $CoverPageArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
          "Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
          "ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
        "Randlinie", "Austin", "Integral", "Facette")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
          "Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
          "Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
        "Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
      }
    }

    'en-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
          "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
          "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
        "Whisp")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
          "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
        "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
      }
    }

    'es-'	{
      if($WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
          "Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
          "Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
        "Semáforo", "Slice (luz)", "Vista principal", "Whisp")
      }
      elseif($WordVersion -eq $wdWord2013) {
        $CoverPageArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
          "Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
          "Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
        "Ion (claro)", "Integral", "Con bandas")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
          "Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
          "Moderno", "Mosaicos", "Movimiento", "Papel periódico",
        "Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
      }
    }

    'fi-'	{
      if($WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Filigraani", "Integraali", "Ioni (tumma)",
          "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
          "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
        "Kuiskaus", "Liike", "Ruudukko", "Sivussa")
      }
      elseif($WordVersion -eq $wdWord2013) {
        $CoverPageArray = ("Filigraani", "Integraali", "Ioni (tumma)",
          "Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
          "Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
        "Kiehkura", "Liike", "Ruudukko", "Sivussa")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
          "Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
          "Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
        "Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
      }
    }

    'fr-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("À bandes", "Austin", "Facette", "Filigrane", 
          "Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
          "Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
        "Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alphabet", "Annuel", "Austère", "Austin", 
          "Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
          "Exposition", "Guide", "Ligne latérale", "Moderne", 
          "Mosaïques", "Mots croisés", "Papier journal", "Perspective",
        "Quadrillage", "Rayures fines", "Transcendant")
      }
    }

    'nb-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Bevegelse", "Dempet", "Fasett", "Filigran",
          "Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
          "Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
        "ViewMaster")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
          "Bevegelse", "Engasjement", "Enkel", "Fliser", "Konservativ",
          "Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
        "Smale striper", "Stabler", "Transcenderende")
      }
    }

    'nl-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
          "Integraal", "Ion (donker)", "Ion (licht)", "Raster",
          "Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
        "Terugblik", "Terzijde", "ViewMaster")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
          "Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
          "Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
          "Puzzel", "Raster", "Stapels",
        "Tegels", "Terzijde")
      }
    }

    'pt-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
          "Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
          "Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
        "Retrospectiva", "Semáforo")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
          "Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
          "Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
        "Quebra-cabeça", "Transcend") 
      }
    }

    'sv-'	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016) {
        $CoverPageArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
          "Jon (mörkt)", "Knippe", "Rutnät", "Rörelse", "Sektor (ljus)", "Sektor (mörk)",
        "Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
      }
      elseif($WordVersion -eq $wdWord2010) {
        $CoverPageArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
          "Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
          "Rörelse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
        "Övergående")
      }
    }

    'zh-'	{
      if($WordVersion -eq $wdWord2010 -or $WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016)
      {
        $CoverPageArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
          '离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
          '切片(深色)', '丝状', '网格', '镶边', '信号灯',
        '运动型')
      }
    }

    Default	{
      if($WordVersion -eq $wdWord2013 -or $WordVersion -eq $wdWord2016)
      {
        $CoverPageArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
          "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
          "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
        "Whisp")
      }
      elseif($WordVersion -eq $wdWord2010)
      {
        $CoverPageArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
          "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
        "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
      }
    }
  }
	
  if($CoverPageArray -contains $CoverPage)
  {
    $CoverPageArray = $Null
    Return $True
  }
  else
  {
    $CoverPageArray = $Null
    Return $False
  }
}

Function Get-WordCultureCode {
  Param(
    [int]$WordValue
  )
	
  #codes obtained from http://support.microsoft.com/kb/221435
  #http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
  $CatalanArray = 1027
  $ChineseArray = 2052,3076,5124,4100
  $DanishArray = 1030
  $DutchArray = 2067, 1043
  $EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
  $FinnishArray = 1035
  $FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
  $GermanArray = 1031, 3079, 5127, 4103, 2055
  $NorwegianArray = 1044, 2068
  $PortugueseArray = 1046, 2070
  $SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
  $SwedishArray = 1053, 2077

  #ca - Catalan
  #da - Danish
  #de - German
  #en - English
  #es - Spanish
  #fi - Finnish
  #fr - French
  #nb - Norwegian
  #nl - Dutch
  #pt - Portuguese
  #sv - Swedish
  #zh - Chinese

  Switch ($WordValue)
  {
    {$CatalanArray -contains $_} {$CultureCode = "ca-"}
    {$ChineseArray -contains $_} {$CultureCode = "zh-"}
    {$DanishArray -contains $_} {$CultureCode = "da-"}
    {$DutchArray -contains $_} {$CultureCode = "nl-"}
    {$EnglishArray -contains $_} {$CultureCode = "en-"}
    {$FinnishArray -contains $_} {$CultureCode = "fi-"}
    {$FrenchArray -contains $_} {$CultureCode = "fr-"}
    {$GermanArray -contains $_} {$CultureCode = "de-"}
    {$NorwegianArray -contains $_} {$CultureCode = "nb-"}
    {$PortugueseArray -contains $_} {$CultureCode = "pt-"}
    {$SpanishArray -contains $_} {$CultureCode = "es-"}
    {$SwedishArray -contains $_} {$CultureCode = "sv-"}
    Default {$CultureCode = "en-"}
  }
	
  Return $CultureCode
}

function Set-SectionTitle {
  param (
    [string]$SectionTitle = '',
    [int]$Style = 3,
    [switch]$NewPage
  )

  if($NewPage) {
    # Insert page break
    $Script:Selection.InsertNewPage()
  }

  Write-WordLine -Style $Style -Tabs 0 -Name $SectionTitle

}

function Get-CompanyName {
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	if($xResult) {
		Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	else {
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		if($xResult) {
			Return Get-LocalRegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		else {
			Return ""
		}
	}
}

function Close-WordDocument {
  
  # Reset Grammar and Spelling options back to their original settings befor closing Word
  $Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
  $Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

  # Pepare file name
  $Script:FileName = ('{0}-{1}.docx' -f $FileName, (Get-Date -f yyyy-MM-dd))
  $Script:FileNameWord = "$($Script:FileName)"

  if($Script:WordVersion -eq $wdWord2010) {
    
    # Set default document type
    $SaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
    
    # Save Word document
    $Script:WordDocument.SaveAs([REF]$Script:FileNameWord, [ref]$SaveFormat)
    
  }
  elseif($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
    # Save as Word Default document
    $Script:WordDocument.SaveAs2([REF]$Script:FileNameWord, [ref]$wdFormatDocumentDefault)
  }

  # Close document
  $Script:WordDocument.Close()

  # Quit Word
  $Script:Word.Quit()

  # Finally, cleanup Word variable 
  [Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
  if(Test-Path variable:global:word) {
    Remove-Variable -Name word -Scope Global 4>$Null
  }
  $SaveFormat = $Null
  [gc]::collect() 
  [gc]::WaitForPendingFinalizers()

}

function Select-WordEndOfDocument {
  # Return focus to main document    
  $Script:WordDocument.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

  # Move to the end of the current document
  $Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

function New-MicrosoftWordDocument {

  # Create a new ComObject instance of Word

  Write-Verbose -Message "$(Get-Date): Create Word ComObject"
  $Script:Word = New-Object -ComObject "Word.Application" -ErrorAction SilentlyContinue 4>$Null
  
  if(!$? -or $Null -eq $Script:Word) {

    # Ooops, something went wrong
    Write-Warning -Message 'The Word ComObject could not be created. You may need to install Word or repair an existing installation.'
    $ErrorActionPreference = $SavedErrerActionPreference

    Exit
  }
  
  # As we have a Word ComObject, we can continue
  # Let's determine the language version 
  if((Get-ValidStateProp -Object $Script:Word -TopLevel Language -SecondLevel Value__ )) {
    [int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
  }
  else {
    [int]$Script:WordLanguageValue = [int]$Script:Word.Language
  }
  
  Write-Verbose ('{0}: Word language value is {1}' -f (Get-Date), $Script:WordLanguageValue)

  $Script:WordCultureCode = Get-WordCultureCode -WordValue $Script:WordLanguageValue
	
  Set-WordHashTable -CultureCode $Script:WordCultureCode

  # Check Word product version
  # Supported versions Word 2010 or newer
  [int]$Script:WordVersion = [int]$Script:Word.Version

  if($Script:WordVersion -eq $wdWord2016) {
    $Script:WordProduct = "Word 2016"
  }
  elseif($Script:WordVersion -eq $wdWord2013)	{
    $Script:WordProduct = "Word 2013"
  }
  elseif($Script:WordVersion -eq $wdWord2010) {
    $Script:WordProduct = "Word 2010"
  }
  elseif($Script:WordVersion -eq $wdWord2007)	{
    $ErrorActionPreference = $SavedErrerActionPreference
    Write-Error -Message "Microsoft Word 2007 is no longer supported.`nScript will end."
    Stop-Script
  }
  else {
    $ErrorActionPreference = $SavedErrerActionPreference
    Write-Error -Message "You are running an untested or unsupported version of Microsoft Word.`nScript will end.`nPlease send info on your version of Word to thomas@mcsmemail.de"
    Stop-Script
  }

  # only validate CompanyName if the field is blank
  if([String]::IsNullOrEmpty($Script:CoName)) {
    Write-Verbose -Message "$(Get-Date): Company name is blank. Retrieve company name from registry."
    
    $TmpName = Get-CompanyName
		
    if([String]::IsNullOrEmpty($TmpName)) {
      Write-Warning -Message 'Company Name is blank so Cover Page will not show a Company Name.'
      Write-Warning -Message 'Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.'
      Write-Warning -Message 'You may want to use the -CompanyName parameter if you need a Company Name on the cover page.'
    }
    else {
      $Script:CoName = $TmpName
      Write-Verbose -Message ('{0}: Updated company name (CoName) to {1}' -f (Get-Date), $Script:CoName)
    }
  }

  # Check Word cover page and set localized template name
  if($Script:WordCultureCode -ne 'en-') {

    Write-Verbose ('{0}: Check Default Cover Page for {1}' -f (Get-Date), $WordCultureCode)

    [bool]$CoverPageChanged = $False
    Switch ($Script:WordCultureCode) {
      'ca-'	{
        if($CoverPage -eq 'Sideline') {
          $CoverPage = 'Línia lateral'
          $CoverPageChanged = $True
        }
      }

      'da-'	{
        if($CoverPage -eq 'Sideline') {
          $CoverPage = 'Sidelinje'
          $CoverPageChanged = $True
        }
      }

      'de-'	{
        if($CoverPage -eq 'Sideline') {
          $CoverPage = 'Randlinie'
          $CoverPageChanged = $True
        }
      }

      'es-'	{
        if($CoverPage -eq "Sideline")	{
          $CoverPage = "Línea lateral"
          $CoverPageChanged = $True
        }
      }

      'fi-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "Sivussa"
          $CoverPageChanged = $True
        }
      }

      'fr-'	{
        if($CoverPage -eq "Sideline")	{
          if($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016) {
            $CoverPage = "Lignes latérales"
            $CoverPageChanged = $True
          }
          else {
            $CoverPage = "Ligne latérale"
            $CoverPageChanged = $True
          }
        }
      }

      'nb-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "Sidelinje"
          $CoverPageChanged = $True
        }
      }

      'nl-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "Terzijde"
          $CoverPageChanged = $True
        }
      }

      'pt-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "Linha Lateral"
          $CoverPageChanged = $True
        }
      }

      'sv-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "Sidlinje"
          $CoverPageChanged = $True
        }
      }

      'zh-'	{
        if($CoverPage -eq "Sideline") {
          $CoverPage = "边线型"
          $CoverPageChanged = $True
        }
      }
    }

    if($CoverPageChanged) {
      Write-Verbose ('{0}: Changed Default Cover Page from "Sideline" to "{1}"' -f (Get-Date), $CoverPage)
    }
  }

  Write-Verbose ('{0}: Validate cover page {1} for culture code {2}' -f (Get-Date), $CoverPage, $Script:WordCultureCode)
	
  [bool]$ValidCoverPage = $False	
  $ValidCoverPage = Test-WordCoverPage -WordVersion $Script:WordVersion -CoverPage $CoverPage -CultureCode $Script:WordCultureCode
  
  if(!$ValidCoverPage)	{

    # stop script, if Word cover page is not valid
    $ErrorActionPreference = $SavedErrerActionPreference
    Write-Verbose -Message ('{0}: Word language value {1}' -f (Get-Date), $Script:WordLanguageValue)
    Write-Verbose -Message ('{0}: Culture code {1}' -f (Get-Date), $Script:WordCultureCode)
    Write-Error -Message ("For {0}, {1} is not a valid Cover Page option.`nScript cannot continue." -f $Script:WordProduct, $CoverPage)

    Stop-Script
  }

  # Show script options
  Show-ScriptOptions

  # Run Word instance invisble
  $Script:Word.Visible = $false

  # http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
  # using Jeff's Demo-WordReport.ps1 file for examples
  Write-Verbose "$(Get-Date): Load Word Templates"

  [bool]$Script:CoverPagesExist = $False
  [bool]$BuildingBlocksExist = $False

  $Script:Word.Templates.LoadBuildingBlocks()

  # Word 2010/2013/2016
  $BuildingBlocksCollection = $Script:Word.Templates | Where-Object {$_.Name -eq 'Built-In Building Blocks.dotx'}
  
  $part = $Null

  # Fetch cover page
  $BuildingBlocksCollection | ForEach-Object{
    if ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) {
      $BuildingBlocks = $_
    }
  } 

  # Check if cover page exists in current Word setup
  if($Null -ne $BuildingBlocks)	{

    $BuildingBlocksExist = $True

    try {
      $part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
    }
    catch	{
      $part = $Null
    }

    if($Null -ne $part) {
      $Script:CoverPagesExist = $True
    }
  }
    
  if(!$Script:CoverPagesExist) {
    Write-Verbose -Message ('Cover Pages are not installed or the Cover Page {1} does not exist.' -f $CoverPage)
    Write-Warning -Message ('Cover Pages are not installed or the Cover Page {0} does not exist.' -f $CoverPage)
    Write-Warning -Message 'This report will not have a Cover Page.'
  }

  # Create a new Word document in the current Word instance
  $Script:WordDocument = $Script:Word.Documents.Add()

  if($Null -eq $Script:WordDocument) {
    # failed to create a new Word document
    		
    $ErrorActionPreference = $SavedErrerActionPreference
    Write-Error -Message 'An empty Word document could not be created. Script cannot continue.'

    Stop-Script
  }

  $Script:Selection = $Script:Word.Selection
  if($Null -eq $Script:Selection) {
    # Some error occured 

    $ErrorActionPreference = $SavedErrerActionPreference
    Write-Error -Message 'An unknown error happened selecting the entire Word document for default formatting options. Script cannot continue.'
		
    Stop-Script
  }

  # set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
  # 36 = .50"
  $Script:Word.ActiveDocument.DefaultTabStop = 36

  # Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
  # Save current options first before turning them off
  $Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
  $Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
  $Script:Word.Options.CheckGrammarAsYouType = $False
  $Script:Word.Options.CheckSpellingAsYouType = $False  

  if($BuildingBlocksExist) {

    # insert new page, getting ready for table of contents
    $part.Insert($Script:Selection.Range,$True) | Out-Null
    $Script:Selection.InsertNewPage()

    # table of contents
    $toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		
    if($Null -eq $toc) {
      Write-Warning -Message 'This report will not have a Table of Contents.'
    }
    else {
      $toc.Insert($Script:Selection.Range,$True) | Out-Null
    }
  }
  else {
    Write-Warning -Message 'Table of Contents (TOC) are not installed so this report will not have a Table of Contents.'
  }  

  #region Footer 

  # Set the document footer text
  [string]$FooterText = ('Report created by {0} on {1}' -f $UserName, $GeneratedOn)

  # Fetch footer for additional configuration
  $Script:WordDocument.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter

  # Get the footer and format font
  $Footers = $Script:WordDocument.Sections.Last.Footers

  ForEach ($Footer in $Footers) {
    if($Footer.exists) {
      # Set font
      $Footer.Range.Font.name = "Calibri"
      $Footer.Range.Font.size = 8
      $Footer.Range.Font.Italic = $True
      $Footer.Range.Font.Bold = $True
    }
  } 

  $Script:Selection.HeaderFooter.Range.Text = $FooterText

  # Add page numbering
  $Script:Selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

  #endregion

  Select-WordEndOfDocument


}

Function Get-ValidStateProp {
  param (
    [object] $Object, 
    [string] $TopLevel, 
    [string] $SecondLevel
  )
  if( $Object ) {
    if((Get-Member -Name $TopLevel -InputObject $Object)) {
    
      if((Get-Member -Name $SecondLevel -InputObject $Object.$TopLevel)) {
        Return $True
      }
    }
  }
  Return $False
}

function Set-DocumentProperty {
  <#
      .SYNOPSIS
      Function to set the Title Page document properties in MS Word
      .DESCRIPTION
      Long description
      .PARAMETER Document
      Current Document Object
      .PARAMETER DocProperty
      Parameter description
      .PARAMETER Value
      Parameter description
      .EXAMPLE
      Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
      .EXAMPLE
      Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
      .EXAMPLE
      Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
      .EXAMPLE
      Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
      .NOTES
      Function Created by Jim Moyle June 2017
      Twitter : @JimMoyle
  #>
  param (
    [object]$Document,
    [String]$DocProperty,
    [string]$Value
  )
  try {
    $binding = "System.Reflection.BindingFlags" -as [type]
    $builtInProperties = $Document.BuiltInDocumentProperties
    $property = [__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
    [__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
  }
  catch {
    Write-Warning -Message ('Failed to set {0} to {1}' -f $DocProperty, $Value)
  }
}

function Update-DocumentProperty {
  param(
    [string]$DocumentTitle = '',
    [string]$AbstractTitle,
    [string]$SubjectTitle = 'Exchange Organization Report',
    [string]$Author = ''
  )

  if($ExportTo -eq 'MSWord') {
    if($Script:CoverPagesExist) {
      
      # Set document properties
      Set-DocumentProperty -Document $Script:WordDocument -DocProperty Author -Value $Author
      Set-DocumentProperty -Document $Script:WordDocument -DocProperty Company -Value $Script:CoName
      Set-DocumentProperty -Document $Script:WordDocument -DocProperty Subject -Value $SubjectTitle
      Set-DocumentProperty -Document $Script:WordDocument -DocProperty Title -Value $DocumentTitle

      # Fetch cover page XML 
      $CoverPageXml = $Script:WordDocument.CustomXMLParts | Where-Object {$_.NamespaceURI -match "coverPageProps$"}

      # Fetch abstract XML part
      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'Abstract'}
			
      #set the text
      if([String]::IsNullOrEmpty($Script:CoName))	{
        [string]$Abstrac = $AbstractTitle
      }
      else {
        [string]$Abstract = ('{0} for {1}' -f $AbstractTitle, $Script:CoName)
      }

      $AbstractXml.Text = $Abstract

      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'CompanyAddress'}
      [string]$AbstractXmlstract = $CompanyAddress
      $AbstractXml.Text = $AbstractXmlstract

      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'CompanyEmail'}
      [string]$AbstractXmlstract = $CompanyEmail
      $AbstractXml.Text = $AbstractXmlstract

      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'CompanyFax'}
      [string]$AbstractXmlstract = $CompanyFax
      $AbstractXml.Text = $AbstractXmlstract

      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'CompanyPhone'}
      [string]$AbstractXmlstract = $CompanyPhone
      $AbstractXml.Text = $AbstractXmlstract

      $AbstractXml = $CoverPageXml.documentelement.ChildNodes | Where-Object {$_.basename -eq 'PublishDate'}
      [string]$AbstractXmlstract = (Get-Date -Format d).ToString()
      $AbstractXml.Text = $AbstractXmlstract

      Write-Verbose "$(Get-Date): Update the Table of Contents"
      # Update the Table of Contents
      $Script:WordDocument.TablesOfContents.item(1).Update()
      $CoverPageXml = $Null
      $AbstractXml = $Null
      $AbstractXmlstract = $Null
    }
  }
}

function Add-WordTable {
  Param	(
    # Array of Hashtable (including table headers)
    [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName='Hashtable', Position=0)]
    [ValidateNotNullOrEmpty()] [Collections.Hashtable[]] $Hashtable,
    # Array of PSCustomObjects
    [Parameter(Mandatory, ValueFromPipelineByPropertyName, ParameterSetName='CustomObject', Position=0)]
    [ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
    # Array of Hashtable key names or PSCustomObject property names to include, in display order.
    # if not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
    [Parameter(ValueFromPipelineByPropertyName)] [AllowNull()] [string[]] $Columns = $Null,
    # Array of custom table header strings in display order.
    [Parameter(ValueFromPipelineByPropertyName)] [AllowNull()] [string[]] $Headers = $Null,
    # AutoFit table behavior.
    [Parameter(ValueFromPipelineByPropertyName)] [AllowNull()] [int] $AutoFit = -1,
    # List view (no headers)
    [Switch] $List,
    # Grid lines
    [Switch] $NoGridLines,
    [Switch] $NoInternalGridLines,
    # Built-in Word table formatting style constant
    # Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
    [Parameter(ValueFromPipelineByPropertyName)] [int] $Format = 0
  )

  Begin {
    ## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
    if(($Null -eq $Columns) -and ($Null -ne $Headers)) {
      Write-Warning -Message 'No columns specified and therefore, specified headers will be ignored.'
      $Columns = $Null
    }
    elseif(($Null -ne $Columns) -and ($Null -ne $Headers)) {
      ## Check if number of specified -Columns matches number of specified -Headers
      if($Columns.Length -ne $Headers.Length) {
        Write-Error -Message 'The specified number of columns does not match the specified number of headers.'
      }
    } ## end elseif
  } ## end Begin

  Process	{
    ## Build the Word table data string to be converted to a range and then a table later.
    [Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder

    Switch ($PSCmdlet.ParameterSetName) {
      'CustomObject'  {
        if($Null -eq $Columns) {
          ## Build the available columns from all available PSCustomObject note properties
          [string[]] $Columns = @()
          ## Add each NoteProperty name to the array
          ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) { 
            $Columns += $Property.Name
          }
        }

        ## Add the table headers from -Headers or -Columns (except when in -List(view)
        if(-not $List) 	{
          if($Null -ne $Headers){
            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers))
          }
          else { 
            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns))
          }
        }

        ## Iterate through each PSCustomObject
        ForEach($Object in $CustomObject) {
          $OrderedValues = @()
          ## Add each row item in the specified order
          ForEach($Column in $Columns) { 
            $OrderedValues += $Object.$Column
          }
          ## Use the ordered list to add each column in specified order
          [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues))
        } ## end ForEach
        Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count))
      } ## end CustomObject

      Default {   
        ## Hashtable
        if($Null -eq $Columns) {
          ## Build the available columns from all available hashtable keys. Hopefully
          ## all Hashtables have the same keys (they should for a table).
          $Columns = $Hashtable[0].Keys
        }

        ## Add the table headers from -Headers or -Columns (except when in -List(view)
        if(-not $List) {
          if($Null -ne $Headers) { 
            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers))
          }
          else {
            [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns))
          }
        }
                
        ## Iterate through each Hashtable
        Write-Debug ("$(Get-Date): `t`tBuilding table rows")
        ForEach($Hash in $Hashtable) {
          $OrderedValues = @()
          ## Add each row item in the specified order
          ForEach($Column in $Columns) { 
            $OrderedValues += $Hash.$Column
          }
          ## Use the ordered list to add each column in specified order
          [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues))
        } ## end ForEach

      } ## end default
    } ## end switch

    ## Create a MS Word range and set its text to our tab-delimited, concatenated string
    Write-Debug ("$(Get-Date): `t`tBuilding table range")
    $WordRange = $Script:WordDocument.Application.Selection.Range
    $WordRange.Text = $WordRangeString.ToString()

    ## Create hash table of named arguments to pass to the ConvertToTable method
    $ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs }

    ## Negative built-in styles are not supported by the ConvertToTable method
    if($Format -ge 0) {
      $ConvertToTableArguments.Add("Format", $Format)
      $ConvertToTableArguments.Add("ApplyBorders", $True)
      $ConvertToTableArguments.Add("ApplyShading", $True)
      $ConvertToTableArguments.Add("ApplyFont", $True)
      $ConvertToTableArguments.Add("ApplyColor", $True)
      if(!$List) { 
        $ConvertToTableArguments.Add("ApplyHeadingRows", $True)
      }
      $ConvertToTableArguments.Add("ApplyLastRow", $True)
      $ConvertToTableArguments.Add("ApplyFirstColumn", $True)
      $ConvertToTableArguments.Add("ApplyLastColumn", $True)
    }

    ## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
    ## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
    ## Store the table reference just in case we need to set alternate row coloring
    $WordTable = $WordRange.GetType().InvokeMember(
      "ConvertToTable",                               # Method name
      [Reflection.BindingFlags]::InvokeMethod, # Flags
      $Null,                                          # Binder
      $WordRange,                                     # Target (self!)
      ([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
      $Null,                                          # Modifiers
      $Null,                                          # Culture
      ([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
    )

    ## Implement grid lines (will wipe out any existing formatting
    if($Format -lt 0) {
      $WordTable.Style = $Format
    }

    ## Set the table autofit behavior
    if($AutoFit -ne -1) { 
      $WordTable.AutoFitBehavior($AutoFit)
    }

    if(!$List) {
      #the next line causes the heading row to flow across page breaks
      $WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue
    }

    if(!$NoGridLines) {
      $WordTable.Borders.InsideLineStyle = $wdLineStyleSingle
      $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle
    }
    if($NoGridLines) {
      $WordTable.Borders.InsideLineStyle = $wdLineStyleNone
      $WordTable.Borders.OutsideLineStyle = $wdLineStyleNone
    }
    if($NoInternalGridLines) {
      $WordTable.Borders.InsideLineStyle = $wdLineStyleNone
      $WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle
    }

    Return $WordTable

  } ## end Process
}

function Set-WordCellFormat {
  [CmdletBinding(DefaultParameterSetName='Collection')]
  Param (
    # Word COM object cell collection reference
    [Parameter(Mandatory, ValueFromPipeline, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
    # Word COM object individual cell reference
    [Parameter(Mandatory, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
    # Hashtable of cell co-ordinates
    [Parameter(Mandatory, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [Collections.Hashtable[]] $Coordinates,
    # Word COM object table reference
    [Parameter(Mandatory, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
    # Font name
    [Parameter()] [AllowNull()] [string] $Font = $Null,
    # Font color
    [Parameter()] [AllowNull()] $Color = $Null,
    # Font size
    [Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
    # Cell background color
    [Parameter()] [AllowNull()] $BackgroundColor = $Null,
    # Force solid background color
    [Switch] $Solid,
    [Switch] $Bold,
    [Switch] $Italic,
    [Switch] $Underline
  )

  Begin {
  }

  Process {
    Switch ($PSCmdlet.ParameterSetName) {
      'Collection' {
        ForEach($Cell in $Collection) 
        {
          if($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor }
          if($Bold) { $Cell.Range.Font.Bold = $true }
          if($Italic) { $Cell.Range.Font.Italic = $true }
          if($Underline) { $Cell.Range.Font.Underline = 1 }
          if($Null -ne $Font) { $Cell.Range.Font.Name = $Font }
          if($Null -ne $Color) { $Cell.Range.Font.Color = $Color }
          if($Size -ne 0) { $Cell.Range.Font.Size = $Size }
          if($Solid) { $Cell.Shading.Texture = 0 } ## wdTextureNone
        } # end ForEach
      } # end Collection
      'Cell' 
      {
        if($Bold) { $Cell.Range.Font.Bold = $true }
        if($Italic) { $Cell.Range.Font.Italic = $true }
        if($Underline) { $Cell.Range.Font.Underline = 1 }
        if($Null -ne $Font) { $Cell.Range.Font.Name = $Font }
        if($Null -ne $Color) { $Cell.Range.Font.Color = $Color }
        if($Size -ne 0) { $Cell.Range.Font.Size = $Size }
        if($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor }
        if($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
      } # end Cell
      'Hashtable' 
      {
        ForEach($Coordinate in $Coordinates) 
        {
          $Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column)
          if($Bold) { $Cell.Range.Font.Bold = $true }
          if($Italic) { $Cell.Range.Font.Italic = $true }
          if($Underline) { $Cell.Range.Font.Underline = 1 }
          if($Null -ne $Font) { $Cell.Range.Font.Name = $Font }
          if($Null -ne $Color) { $Cell.Range.Font.Color = $Color }
          if($Size -ne 0) { $Cell.Range.Font.Size = $Size }
          if($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor }
          if($Solid) { $Cell.Shading.Texture = 0 } ## wdTextureNone
        }
      } # end Hashtable
    } # end switch
  } # end process
}

function Write-EmptyWordLine {
  # always default back to default size 
  Write-WordLine -Style 0 -Tabs 0 -Name '' -FontSize $WordDefaultFontSize
}

#endregion

#region General functions

function Show-ScriptOptions {

  if($ShowScriptOptions) {

    if($ExportTo -eq 'MSWord') {
      Write-Verbose -Message ('Company Name    : {0}' -f $Script:CoName)
    }
  }
}

function Get-BasicScriptInformation {
}

function Test-ExchangeManagementShellVersion {
  if ((Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction SilentlyContinue)) {
  
    $E2010 = $false
    
    if (Get-ExchangeServer | Where-Object {$_.AdminDisplayVersion.Major -gt 14}) {
      Write-Warning -Message "Exchange 2010 or higher detected. You'll get better results if you run this script from the latest management shell"
    }
  }
  else{
    
    $E2010 = $true

    $localversion = $localserver = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiProductMajor

    if ($localversion -eq 15) { $E2013 = $true }
  }
}

function Optimize-String {
  param (
    [string]$Value
  )

  $Value = $Value -replace "`r`n", "`n"

  $Value = $Value -replace "`n", " "

  $Value
}

#endregion

#region Active Directory Stuff
function Get-ActiveDirectoryInformation {

  # Get forest information
  $Script:Forest = Get-ADForest -Identity $ADForest -ErrorAction 0

  $Script:Domains = $Script:Forest.Domains | Sort-Object 
  $Script:ConfigNC = (Get-ADRootDSE -Server $ADForest -ErrorAction 0).ConfigurationNamingContext

}

#endregion

#region Exchange Organization 

function Get-ExchangeOrganizationConfig {

  Show-ProgressBar -Status 'Get-ExchangeOrganization' -PercentComplete 10 -Stage 1

  $OrgConfig = Get-OrganizationConfig 

  $Script:ExchangeOrgName = $OrgConfig.Name

  [Collections.Hashtable[]]$PublicFolderInformation = @()
  
  $PublicFolderInformation += @{ Data = "PublicFolders Enabled"; Value = $OrgConfig.PublicFoldersEnabled; }
  $PublicFolderInformation += @{ Data = "PublicFolders Locked For Migration"; Value = $OrgConfig.PublicFoldersLockedForMigration; }
  $PublicFolderInformation += @{ Data = "PublicFolder Migration Complete"; Value = $OrgConfig.PublicFolderMigrationComplete; }
  $PublicFolderInformation += @{ Data = "PublicFolder Mailboxes Locked For New Connections"; Value = $OrgConfig.PublicFolderMailboxesLockedForNewConnections; }
  $PublicFolderInformation += @{ Data = "PublicFolder Mailboxes Migration Complete"; Value = $OrgConfig.PublicFolderMailboxesMigrationComplete; }
  $PublicFolderInformation += @{ Data = "PublicFolder Show Client Control"; Value = $OrgConfig.PublicFolderShowClientControl; }
  
  # Default 
  $PublicFolderInformation += @{ Data = "Default PublicFolder Age Limit"; Value = $OrgConfig.DefaultPublicFolderAgeLimit; }
  $PublicFolderInformation += @{ Data = "Default PublicFolder IssueWarning Quota"; Value = $OrgConfig.DefaultPublicFolderIssueWarningQuota; }
  $PublicFolderInformation += @{ Data = "Default PublicFolder ProhibitPost Quota"; Value = $OrgConfig.DefaultPublicFolderProhibitPostQuota; }
  $PublicFolderInformation += @{ Data = "Default PublicFolder MaxItemSize"; Value = $OrgConfig.DefaultPublicFolderMaxItemSize; }
  $PublicFolderInformation += @{ Data = "Default PublicFolder DeletedItem Retention"; Value = $OrgConfig.DefaultPublicFolderDeletedItemRetention; }
  $PublicFolderInformation += @{ Data = "Default PublicFolder MovedItem Retention"; Value = $OrgConfig.DefaultPublicFolderMovedItemRetention; }

  $SectionTitle = ('Exchange Organization {0}' -f $OrgConfig.Name)
  
  $Text = ('Active Directory Forest {1} contains an Exchange Organization named {0}. The following table shows selected organization configuration settings active on {2}.' -f $OrgConfig.Name, $Script:Forest.Name, $GeneratedOn)
  
  if($ExportTo -eq 'MSWord') {
    
    Write-WordLine -Style 0 -Tabs 0 -Name $Text 
    Write-EmptyWordLine
    Write-WordLine -Style 2 -Tabs 0 -Name 'Exchange Organization Configuration' 

    $Table = Add-WordTable -Hashtable $PublicFolderInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }

}

function Get-RecipientInformation {

  Show-ProgressBar -Status 'Get Recipient Information - Fetching Mailbox Users' -PercentComplete 15 -Stage 1

  $Mailboxes = Get-Mailbox -Resultsize Unlimited 

  Show-ProgressBar -Status 'Get Recipient Information - Analyzing Mailbox Users' -PercentComplete 15 -Stage 1

  $MailboxCount = ($Mailboxes).Count
  $UserMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'UserMailbox'}).Count
  $RemoteUserMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'RemoteUserMailbox'}).Count

  $SharedMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'SharedMailbox'}).Count
  $RemoteSharedMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'RemoteSharedMailbox'}).Count

  $RoomMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'RoomMailbox'}).Count
  $RemoteRoomMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'RemoteRoomMailbox'}).Count

  $EquipmentMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'EquipmentMailbox'}).Count
  $RemoteEquipmentMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'RemoteEquipmentMailbox'}).Count

  $LinkedMailboxCount = ($Mailboxes | Where-Object{$_.RecipientTypeDetails -eq 'LinkedMailbox'}).Count

  Show-ProgressBar -Status 'Get Recipient Information - Fetching Public Folder Mailboxes' -PercentComplete 15 -Stage 1

  $PublicFolderMailboxCount = (Get-Mailbox -ResultSize Unlimited -PublicFolder).Count

  Show-ProgressBar -Status 'Get Recipient Information - Fetching Arbitration Mailboxes' -PercentComplete 15 -Stage 1

  $ArbitrationMailboxCount = (Get-Mailbox -ResultSize Unlimited -Arbitration).Count

  Show-ProgressBar -Status 'Get Recipient Information - Fetching Mail Contacts' -PercentComplete 15 -Stage 1

  $MailContactCount = (Get-MailContact -ResultSize Unlimited).Count

  Show-ProgressBar -Status 'Get Recipient Information - Fetching Distribution Groups' -PercentComplete 15 -Stage 1

  $DistributionGroups = Get-DistributionGroup -ResultSize Unlimited

  Show-ProgressBar -Status 'Get Recipient Information - Analyzing Distribution Groups' -PercentComplete 15 -Stage 1

  $DistributionGroupCount = ($DistributionGroups | Measure-Object).Count
  $MailUniversalSecurityGroupCount = ($DistributionGroups | Where-Object{$_.RecipientTypeDetails -eq 'MailUniversalSecurityGroup'}).Count
  $DynamicDistributionGroupCount = ($DistributionGroups | Where-Object{$_.RecipientTypeDetails -eq 'DynamicDistributionGroup'}).Count
  $MailUniversalDistributionGroup = ($DistributionGroups | Where-Object{$_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup'}).Count

  [Collections.Hashtable[]]$RecipientInformation = @()

  # Mailboxes and Contacts
  $RecipientInformation += @{ Data = "Mailboxes"; Value = $MailboxCount; }
  $RecipientInformation += @{ Data = "User Mailboxes"; Value = $UserMailboxCount; }
  $RecipientInformation += @{ Data = "Shared Mailboxes"; Value = $SharedMailboxCount; }
  $RecipientInformation += @{ Data = "Room Mailboxes"; Value = $RoomMailboxCount; }
  $RecipientInformation += @{ Data = "Equipment Mailboxes"; Value = $EquipmentMailboxCount; }
  $RecipientInformation += @{ Data = "Linked Mailboxes"; Value = $LinkedMailboxCount; }
  $RecipientInformation += @{ Data = "Public Folder Mailboxes"; Value = $PublicFolderMailboxCount; }
  $RecipientInformation += @{ Data = "Arbitration Mailboxes"; Value = $ArbitrationMailboxCount; }
  $RecipientInformation += @{ Data = "Mail Contacts"; Value = $MailContactCount; }
  
  # Remote Mailboxes
  $RecipientInformation += @{ Data = "Remote User Mailboxes"; Value = $RemoteUserMailboxCount; }
  $RecipientInformation += @{ Data = "Remote Shared Mailboxes"; Value = $RemoteSharedMailboxCount; }
  $RecipientInformation += @{ Data = "Remote Room Mailboxes"; Value = $RemoteRoomMailboxCount; }
  $RecipientInformation += @{ Data = "Remote Equipment Mailboxes"; Value = $RemoteEquipmentMailboxCount; }

  # Groups
  $RecipientInformation += @{ Data = "Distribution Groups"; Value = $DistributionGroupCount; }
  $RecipientInformation += @{ Data = "Dynamic Distribution Groups"; Value = $DynamicDistributionGroupCount; }
  $RecipientInformation += @{ Data = "Mail Universal Distribution Groups"; Value = $MailUniversalDistributionGroup; }
  $RecipientInformation += @{ Data = "Mail Universal Security Groups"; Value = $MailUniversalSecurityGroupCount; }

  $SectionTitle = 'Exchange Recipients'

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $RecipientInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

  }

}

function Get-AcceptedDomainInformation {

  Show-ProgressBar -Status 'Get Accepted Domain Information' -PercentComplete 15 -Stage 1

  $Domains = Get-AcceptedDomain | Sort-Object DomainName 

  [Collections.Hashtable[]]$HashTable = @()

  foreach($Domain in $Domains) {

    $HashTable += @{ 
      DomainName = $Domain.DomainName; 
      Name = $Domain.Name; 
      DomainType = $Domain.DomainType;
      IsDefault = $Domain.Default;
    }
  }

  $SectionTitle = 'Accepted Domains'
  $Text = ('The Exchange Organization contains {0} accepted domains.' -f $Domains.Count)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 
    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns DomainName, Name, DomainType, IsDefault `
    -Headers 'Domain Name','Name','Domain Type','Default' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

    $Table.Columns.Item(1).Width = 150
    $Table.Columns.Item(2).Width = 150
    $Table.Columns.Item(3).Width = 100
    $Table.Columns.Item(4).Width = 50

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument
    $Table = $Null
    Write-EmptyWordLine
  }
}

function Expand-Object {
  param(
    [psobject]$Object
  )

  [Collections.Hashtable[]]$ObjectInformation = @()

  $Object | Get-Member -Type property | foreach name | foreach {

    $Value = ''

    if($Object.$_ -eq $null) {
      $Value = ''
    }
    else { 
      $Value = $Object.$_.ToString()
    }

    try { 

      if($Object.$_.GetType().BaseType.FullName -eq 'Microsoft.Exchange.Data.MultiValuedPropertyBase' `
      -or $Object.$_.GetType().Name -eq 'ApprovedApplicationCollection') {
        $Value = ($Object.$_ -join ', ')
      }
      # elseif ($Object.$_.GetType().BaseType.GenericTypeArguments.Name -eq 'ADObjectId') {
      elseif($Object.$_ -eq 'RoleAssignments') { 
        #$Value = 'Not exported to Word'
      }
      elseif($Object.$_.GetType().FullName -eq 'Microsoft.Exchange.MessagingPolicies.Rules.Tasks.TransportRuleAction[]') {
        $Value = $NotExported
      }
      elseif($Object.$_.GetType().FullName -eq 'Microsoft.Exchange.MessagingPolicies.Rules.Tasks.TransportRulePredicate[]') {
        $Value = $NotExported
      }
      elseif($Object.$_.GetType().FullName -eq 'Microsoft.Exchange.Configuration.Tasks.RecipientIdParameter[]') {
        $Value = $NotExported 
      }
      elseif($Object.$_.GetType().BaseType.GenericTypeArguments.Name -eq 'ADObjectId') { 
        $Value = ($Object.$_.Name -join ', ')
      }
    }
    catch {
      $Value = "Error while converting $($Object.Name)"    }

    # add to hash table
    #$Value = Optimize-String -Value $Value

    $ObjectInformation += @{ Data = "$($_)"; Value = $($Value); } 

    # Write-Verbose "$($_) : $($Object.$_.GetType().ToString())"
  }

  $ObjectInformation

}

function Get-TransportConfigInformation {

  Show-ProgressBar -Status 'Get Transport Config Information' -PercentComplete 15 -Stage 1

  $TransportConfig = Get-TransportConfig

  $SomeCount = 0

  [Collections.Hashtable[]]$ObjectInformation = @()

  $ObjectInformation = Expand-Object -Object $TransportConfig

  if($TransportConfig.ExternalPostmasterAddress -eq '') {
    $ErrorText += 'The ExternalPostmasterAddress is not configured. According to RFC5321 any system supporting mail relaying or delivery must support the reserved mailbox with postmaster as a local name.'
  }

  $Script:SafetyNetHoldTime = $ObjectInformation.SafetyNetHoldTime

  $SectionTitle = 'Transport Config'
  
  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle
    #Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Font $WordSmallFontSize
    Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Font $WordSmallFontSize

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null
    
    if($ErrorText -ne '') {
      Write-WordLine -Style 0 -Tabs 0 -Name $ErrorText
    }

    Write-EmptyWordLine
  }
}

function Get-DatabaseOverviewInformation {

  Show-ProgressBar -Status 'Database Overview Information' -PercentComplete 15 -Stage 1

  # Internal Notes
  # When you set the replay lag time you will see a warning about the SafetyNetHoldTime as well. It is always recommended to set Safety Net hold time to the same value or greater value than the replay lag time.
  # https://practical365.com/exchange-server/exchange-server-2013-lagged-database-copies-action/

  $ModernMailboxDatabases = Get-MailboxDatabase -Status | Sort-Object -Property Name
  $ModernMailboxDatabaseCount = ($ModernMailboxDatabases | Measure-Object).Count

  $MailboxDatabases = Get-MailboxDatabase -Status -IncludePreExchange2013 | Sort-Object -Property Name
  $MailboxDatabaseCount = (($MailboxDatabases | Measure-Object).Count) - $ModernMailboxDatabaseCount
  $MailboxDatabasesWithoutCopies = ($MailboxDatabases | ?{$_.ReplicationType -eq 'None'} | Measure-Object).Count
  $RecoveryMailboxCount = ($MailboxDatabases | ?{$_.Recovery -eq $true} | Measure-Object).Count

  $PublicFolderDatabases = Get-PublicFolderDatabase -Status | Sort-Object -Property Name
  $PublicFolderDatabaseCount = ($PublicFolderDatabases | Measure-Object).Count

  [Collections.Hashtable[]]$ObjectInformation = @()

  $ObjectInformation += @{ Data = 'Mailbox Databases'; Value = $ModernMailboxDatabaseCount; }
  $ObjectInformation += @{ Data = 'Single Mailbox Databases'; Value = $MailboxDatabasesWithoutCopies; }
  $ObjectInformation += @{ Data = 'Recovery Mailbox Databases'; Value = $RecoveryMailboxCount; }
  $ObjectInformation += @{ Data = 'Legacy Mailbox Databases'; Value = $MailboxDatabaseCount; }
  $ObjectInformation += @{ Data = 'Legacy Public Folder Databases'; Value = $PublicFolderDatabaseCount; }

  $SectionTitle = 'Database Overview'
  $Text = ''

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 
    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

function Get-AdminPermissionInformation {

  Show-ProgressBar -Status 'Get Permission Information' -PercentComplete 15 -Stage 1

  $RoleGroups = Get-RoleGroup | Sort-Object Name 

  $RoleCount = ($RoleGroups | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach($RoleGroup in $RoleGroups) {
    
    $RoleGroupMembers = Get-RoleGroup $RoleGroup.Name | Get-RoleGroupMember

    if($RoleGroupMembers.Count -ne 0) { 

      $Members = [Collections.ArrayList]@()

      foreach($Member in $RoleGroupMembers) {

        $Members.Add(('{0} ({1})' -f $Member.Name, $Member.RecipientType)) | Out-Null
      }

      $MembersString = $Members -join ', '
    }
    else {
      $MembersString = 'None'
    }
  
    $HashTable += @{ 
      RoleGroup = $RoleGroup.Name; 
      RoleGroupMembers = $MembersString; 
    } 
    
  }

  $Text = "The Exchange Organization has $($RoleCount) administrative role groups."

  # Write to Word
  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name 'Admin Roles' 
    
    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns RoleGroup, RoleGroupMembers `
    -Headers 'Role Group Name','Role Group Members' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 150
    $Table.Columns.Item(2).Width = 250

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

function Get-UserRoleAssignmentPolicies {

  Show-ProgressBar -Status 'Get User Role Assignment Policies' -PercentComplete 15 -Stage 1

  $RoleAssignmentPolicies = Get-RoleAssignmentPolicy | Sort-Object Name

  $PolicyCount = ($RoleAssignmentPolicies | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  [Collections.Hashtable[]]$ObjectInformation = @()

  foreach ($Policy in $RoleAssignmentPolicies) {
    
    Show-ProgressBar -Status ('Fetching data for Role Assignment Policy [{0}]' -f $Policy.Name) -PercentComplete 15 -Stage 1

    $MailboxCount = (Get-Mailbox -ResultSize Unlimited | Where-Object{$_.RoleAssignmentPolicy -eq $Policy.Name}).Count
  
    $HashTable += @{ 
      PolicyName = $Policy.Name; 
      IsDefault = $Policy.IsDefault;
      AssignedUsers = $MailboxCount;
    }

  }

  $SectionTitle = 'User Role Assignment Policies'

  $Text = "The Exchange Organization has $($PolicyCount) user role assignment policies."

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns PolicyName, IsDefault, AssignedUsers `
    -Headers 'Policy Name','IsDefault', 'Assigned Users' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 200
    $Table.Columns.Item(2).Width = 80

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 
    
      Write-WordLine -Style 3 -Tabs 0 -Name 'Policy Details'

      foreach ($Policy in $RoleAssignmentPolicies) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Identity)

        $ObjectInformation = Expand-Object -Object $Policy

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
      }
    }
  }
}

function Get-OutlookWebAppPolicies {

  Show-ProgressBar -Status 'Get OWA Mailbox policies' -PercentComplete 15 -Stage 1

  $OwaMailboxPolicies = Get-OwaMailboxPolicy 

  $OwaMailboxPolicyCount = ($OwaMailboxPolicies |Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  [Collections.Hashtable[]]$ObjectInformation = @()

  $CASMailboxes = Get-CASMailbox -ResultSize Unlimited

  foreach ($Policy in $OwaMailboxPolicies) {
    
    Show-ProgressBar -Status ('Fetching data for OWA Policy [{0}]' -f $Policy.Name) -PercentComplete 15 -Stage 1

    $MailboxCount = ($CASMailboxes | Where-Object{$_.OwaMailboxPolicy -eq $Policy.Name}).Count
  
    $HashTable += @{ 
      PolicyName = $Policy.Name; 
      IsDefault = $Policy.IsDefault;
      AssignedUsers = $MailboxCount;
    }

  }

  $SectionTitle = 'OWA Mailbox Policies'

  $Text = ('The Exchange Organization has {0} configured OWA mailbox policies.' -f ($OwaMailboxPolicyCount))

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns PolicyName, IsDefault, AssignedUsers `
    -Headers 'Policy Name','IsDefault', 'Assigned Users' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 200
    $Table.Columns.Item(2).Width = 80

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 
    
      Write-WordLine -Style 3 -Tabs 0 -Name 'Policy Details'

      foreach ($Policy in $OwaMailboxPolicies) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Identity)

        $ObjectInformation = Expand-Object -Object $Policy

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 

      }
    }
  }
}

Function Get-ComplianceInformation {

  Show-ProgressBar -Status 'Compliance Information' -PercentComplete 15 -Stage 1

  $DlpPolicies = Get-DlpPolicy

  $DlpPolicyCount = ($DlpPolicies | Measure-Object).Count

  $Text = ('The Exchange Organization contains {0} data loss prevention (DLP) policies.' -f ($DlpPolicyCount))

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name 'Data Loss Prevention '

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

  }
}

function Get-RetentionPolicyInformation {

  Show-ProgressBar -Status 'Compliance Management - Get Retention Policy Information' -PercentComplete 15 -Stage 1

  $RetentionPolicies = Get-RetentionPolicy | Sort-Object Id

  $Count = ($RetentionPolicies | Measure-Object).Count

  [Collections.Hashtable[]]$ObjectInformation = @()

  foreach($Policy in $RetentionPolicies) { 

    $Tags = $Policy.RetentionPolicyTagLinks.Name -join ', '

    $ObjectInformation += @{ 
      Name = $Policy.Name;
      Tags = $Tags;
      IsDefault = $Policy.IsDefault;
    }
  }

  $SectionTitle = 'Retention Policies'

  $Text = ('The Exchange Organization contains {0} retention policies.' -f $Count)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $ObjectInformation `
    -Columns Name, Tags, IsDefault `
    -Headers 'Name','Retention Policy Tags', 'IsDefault' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 200
    $Table.Columns.Item(3).Width = 80

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

function Get-RetentionPolicyTagInformation {

  Show-ProgressBar -Status 'Compliance Management - Get Retention Policy Tags' -PercentComplete 15 -Stage 1

  $Object = Get-RetentionPolicyTag | Sort-Object Name

  $Count = ($Object | Measure-Object).Count

  [Collections.Hashtable[]]$ObjectInformation = @()

  foreach($PolicyTag in $Object) { 
    $ObjectInformation += @{ 
      Name = $PolicyTag.Name; 
      Type = $PolicyTag.Type.ToString(); 
      AgeLimitForRetention = $PolicyTag.AgeLimitForRetention;
    }
  }

  $SectionTitle = 'Retention Policy Tags'

  $Text = ('The Exchange Organization contains {0} rention policy tags.' -f $Count)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    if($Count -gt 0) { 

      $Table = Add-WordTable -Hashtable $ObjectInformation `
      -Columns Name, Tags, AgeLimitForRetention `
      -Headers 'Name','Type', 'Retention Period' `
      -AutoFit $wdAutoFitFixed

      Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

      $Table.Columns.Item(1).Width = 180
      $Table.Columns.Item(2).Width = 150
      $Table.Columns.Item(3).Width = 80

      $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

      Select-WordEndOfDocument

      $Table = $Null

      Write-EmptyWordLine

      if($IncludeDetails) { 

        if($Count -ne 0) { 
    
          Write-WordLine -Style 3 -Tabs 0 -Name 'Policy Details'

          foreach ($Policy in $Object) {

            Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Identity)

            $ObjectInformation = Expand-Object -Object $Policy

            $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

            Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
            Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

            $Table.Columns.Item(1).Width = 180
            $Table.Columns.Item(2).Width = 300
            $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

            Select-WordEndOfDocument

            $Table = $Null

            Write-EmptyWordLine 
          }
        }
      }
    }
  }
}

function Get-JournalingInformation {

  Show-ProgressBar -Status 'Compliance Management - Get Journaling Rules' -PercentComplete 15 -Stage 1

  $Object = Get-JournalRule | Sort-Object Name 

  $Count = $($Object | Measure-Object).Count

  [Collections.Hashtable[]]$ObjectInformation = @()

  $ObjectInformation += @{ Data = "Mailboxes"; Value = $Count; }

  $SectionTitle = 'Journal Rules'
  $Text = ('The Exchange Organization contains {0} journal rules.' -f $Count)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }


}

function Get-MobileDeviceInformation {

  Show-ProgressBar -Status 'Get Mobile Device Information' -PercentComplete 15 -Stage 1

  $MobileDevices = Get-MobileDevice -Resultsize Unlimited 

  $MobileDeviceCount = ($MobileDevices | Measure-Object).Count
  $MobileDeviceAvtivatedCount = (($MobileDevices| Where-Object{$_.DeviceAccessState -eq 'Allowed'}) | Measure-Object).Count
  $MobileDeviceQuarantinedCount = (($MobileDevices| Where-Object{$_.DeviceAccessState -eq 'Quarantined'}) | Measure-Object).Count
  $MobileDeviceBlockedCount = (($MobileDevices| Where-Object{$_.DeviceAccessState -eq 'Blocked'}) | Measure-Object).Count
  $MobileDeviceDeviceDiscoveryCount = (($MobileDevices| Where-Object{$_.DeviceAccessState -eq 'DeviceDiscovery'}) | Measure-Object).Count
  $MobileDeviceUnknownCount = (($MobileDevices| Where-Object{$_.DeviceAccessState -eq 'Unknown'}) | Measure-Object).Count

  [Collections.Hashtable[]]$ObjectInformation = @()
  
  $ObjectInformation += @{ Data = "Mobile Devices"; Value = $MobileDeviceCount; }
  $ObjectInformation += @{ Data = "Activated Mobile Devices"; Value = $MobileDeviceAvtivatedCount; }
  $ObjectInformation += @{ Data = "Quarantined Mobile Devices"; Value = $MobileDeviceQuarantinedCount; }
  $ObjectInformation += @{ Data = "Blocked Mobile Devices"; Value = $MobileDeviceBlockedCount; }
  $ObjectInformation += @{ Data = "Device Discovery Mobile Devices"; Value = $MobileDeviceDeviceDiscoveryCount; }
  $ObjectInformation += @{ Data = "Unknown Mobile Devices"; Value = $MobileDeviceUnknownCount; }

  [Collections.Hashtable[]]$MobileDeviceTypeTableRowHash = @()
  [Collections.Hashtable[]]$MobileDeviceModelTableRowHash = @()

  $MobileDeviceTypes = $MobileDevices | Group-Object DeviceType | Sort-Object Name
  $MobileDeviceModels = $MobileDevices | Group-Object DeviceModel | Sort-Object Name

  foreach ($Entry in $MobileDeviceTypes) {
  
    $MobileDeviceTypeTableRowHash += @{ 
      DeviceType = $Entry.Name; 
      DeviceCount = $Entry.Count;
    }
  }

  foreach ($Entry in $MobileDeviceModels) {
  
    $MobileDeviceModelTableRowHash += @{ 
      DeviceType = $Entry.Name; 
      DeviceCount = $Entry.Count;
    }
  }

  $SectionTitle = 'Mobile Devices'

  $Text = ''

  if($ExportTo -eq 'MSWord') {
   
    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
    Write-WordLine -Style 0 -Tabs 0 -Name 'Mobile Device Types'

    $Table = Add-WordTable -Hashtable $MObileDeviceTypeTableRowHash `
    -Columns DeviceType, DeviceCount `
    -Headers 'Device Type','Device Count' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

    $Table.Columns.Item(1).Width = 200
    $Table.Columns.Item(2).Width = 80
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
    Write-WordLine -Style 0 -Tabs 0 -Name 'Mobile Device Models'

    $Table = Add-WordTable -Hashtable $MobileDeviceModelTableRowHash `
    -Columns DeviceType, DeviceCount `
    -Headers 'Device Type','Device Count' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 200
    $Table.Columns.Item(2).Width = 80
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

  }
}

function Get-MobileDevicePolicies {

  Show-ProgressBar -Status 'Get Mobile Device Mailbox Policies' -PercentComplete 15 -Stage 1

  $MobileDeviceMailboxPolicies = Get-MobileDeviceMailboxPolicy | Sort-Object Name

  $PolicyCount = ($MobileDeviceMailboxPolicies | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  [Collections.Hashtable[]]$ObjectInformation = @()

  $CASMailboxes = Get-CASMailbox -ResultSize Unlimited

  foreach ($Policy in $MobileDeviceMailboxPolicies) {
    
    Show-ProgressBar -Status ('Fetching data for Mobile Device Policy [{0}]' -f $Policy.Name) -PercentComplete 15 -Stage 1

    $MailboxCount = ($CASMailboxes | Where-Object{$_.ActiveSyncMailboxPolicy -eq $Policy.Name}).Count
  
    $HashTable += @{ 
      PolicyName = $Policy.Name; 
      IsDefault = $Policy.IsDefault;
      AssignedUsers = $MailboxCount;
    }
  }

  $SectionTitle = 'Mobile Device Policies'

  $Text = "The Exchange Organization has $($PolicyCount) mobile device policies."

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle 

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns PolicyName, IsDefault, AssignedUsers `
    -Headers 'Policy Name','IsDefault', 'Assigned User Mailboxes' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 200
    $Table.Columns.Item(2).Width = 80

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 
        
      Write-WordLine -Style 3 -Tabs 0 -Name 'Policy Details'
    
      foreach ($Policy in $MobileDeviceMailboxPolicies) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Name)

        $ObjectInformation = Expand-Object -Object $Policy

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}


function Get-SharingInformation {

  Show-ProgressBar -Status 'Organization Information - Sharing' -PercentComplete 15 -Stage 1

  $Policies = Get-SharingPolicy | Sort-Object Name

  $PolicyCount = ($Policies | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Policy in $Policies) {
  
    $HashTable += @{ 
      PolicyName = $Policy.Name; 
      Domains = $Policy.Domains;
      Default = $Policy.Default;
      Enabled = $Policy.Enabled;
    }
  }

  $SectionTitle = 'Sharing Policy'

  $PolicyText = 'policies'
  if($PolicyCount -eq 1) {$PolicyText = 'policy'}
  $Text = ('The Exchange Organization contains {0} sharing {1}.' -f $PolicyCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns PolicyName, Domains, Default, Enabled `
    -Headers 'Policy Name', 'Domains', 'Default', 'Enabled' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 150
    $Table.Columns.Item(2).Width = 150
    $Table.Columns.Item(3).Width = 70
    $Table.Columns.Item(4).Width = 70

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

  }
}

function Get-AppInformation {

  Show-ProgressBar -Status 'Organization Information - Apps' -PercentComplete 15 -Stage 1

  $Apps = Get-App -OrganizationApp:$true | Sort-Object DisplayName

  $AppCount = ($Apps| Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($App in $Apps) {
  
    $HashTable += @{ 
      DisplayName = $App.DisplayName; 
      ProviderName = $App.ProviderName;
      ProvidedTo = $App.ProvidedTo;
      Enabled = $App.Enabled;
      DefaultStateForUser = $App.DefaultStateForUser;
    }
  }

  $SectionTitle = 'Apps'

  $AppText = 'apps'
  if($AppCount -eq 1) {$PolicyText = 'app'}

  $Text = ('The Exchange Organization contains {0} {1}.' -f $AppCount, $AppText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns DisplayName, ProviderName, ProvidedTo, Enabled, DefaultStateForUser `
    -Headers 'Display Name', 'Provider Name', 'Provided To', 'Enabled', 'Default State For User' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 130
    $Table.Columns.Item(2).Width = 130
    $Table.Columns.Item(3).Width = 70
    $Table.Columns.Item(4).Width = 70
    $Table.Columns.Item(5).Width = 70

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

  }
}

function Get-AddressListInformation {

  Show-ProgressBar -Status 'Organization Information - Address Lists' -PercentComplete 15 -Stage 1

  $AddressLists = Get-AddressList | Sort-Object Name

  $AddressListCount = ($AddressLists| Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($AddressList in $AddressLists) {
  
    $HashTable += @{ 
      Name = $AddressList.Name; 
      DisplayName = $AddressList.DisplayName; 
      RecipientFilter = $AddressList.RecipientFilter; 
    }
  }

  $SectionTitle = 'Address Lists'

  $AppText = 'address lists'
  if($AddressListCount-eq 1) {$PolicyText = 'address list'}

  $Text = ('The Exchange Organization contains {0} {1}.' -f $AddressListCount, $AppText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, DisplayName, RecipientFilter `
    -Headers 'Name', 'Display Name', 'Recipient Filter' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 130
    $Table.Columns.Item(2).Width = 130
    $Table.Columns.Item(3).Width = 180
    

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Address List Details'
    
      foreach ($AddressList in $AddressLists) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $AddressList.Name)

        # store policy details in hash table
        $ObjectInformation = Expand-Object -Object $AddressList

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        # set font to 8pt
        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}

function Get-MalwareInformation {

  Show-ProgressBar -Status 'Protection - Malware' -PercentComplete 15 -Stage 1

  $MalwarePolicies = Get-MalwareFilterPolicy | Sort-Object Name

  $MalwarePolicyCount = ($MalwarePolicies| Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Policy in $MalwarePolicies) {
  
    $HashTable += @{ 
      Name = $Policy.Name; 
      Action = $Policy.Action; 
      CustomNotifications = $Policy.CustomNotifications; 
      IsDefault = $Policy.IsDefault; 
    }
  }

  $SectionTitle = 'Malware Policy'

  $PolicyText = 'malware policies'
  if($MalwarePolicyCount-eq 1) {$PolicyText = 'malware policy'}

  $Text = ('The Exchange Organization contains {0} {1}.' -f $MalwarePolicyCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Action, CustomNotifications, IsDefault `
    -Headers 'Name', 'Action', 'Custom Notifications', 'IsDefault' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 130
    $Table.Columns.Item(2).Width = 130
    $Table.Columns.Item(3).Width = 80
    $Table.Columns.Item(4).Width = 80
    

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Malware Policy Details'

      foreach ($Policy in $MalwarePolicies) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Name)

        $ObjectInformation = Expand-Object -Object $Policy

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }

    }
  }
}

function Get-TransportRuleInformation {

  Show-ProgressBar -Status 'Mail Flow - Transport Rules' -PercentComplete 15 -Stage 1

  $TransportRules = Get-TransportRule | Sort-Object Priority

  $TransportRuleCount = ($TransportRules| Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Rule in $TransportRules) {
  
    $HashTable += @{ 
      Name = $Rule.Name; 
      Priority = $Rule.Priority; 
      State = $Rule.State; 
      Mode = $Rule.Mode; 
      Comments = (Optimize-String -Value $Rule.Comments); 
    }
  }

  $SectionTitle = 'Transport Rules'

  $PolicyText = 'rules'
  if($TransportRuleCount-eq 1) {$PolicyText = 'rule'}

  $Text = ('The Exchange Organization contains {0} transport {1}.' -f $TransportRuleCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Priority, State, Mode, Comments `
    -Headers 'Name', 'Priority', 'State', 'Mode', 'Comments' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 140
    $Table.Columns.Item(2).Width = 50
    $Table.Columns.Item(3).Width = 50
    $Table.Columns.Item(4).Width = 50
    $Table.Columns.Item(5).Width = 140
    

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Transport Rule Details'

      foreach ($Rule in $TransportRules) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Rule.Name)

        $ObjectInformation = Expand-Object -Object $Rule

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }

    }
  }
}

function Get-EmailAddressPolicyInformation {

  Show-ProgressBar -Status 'Mail Flow - Email Address Policies' -PercentComplete 15 -Stage 1

  $EmailAddressPolicies = Get-EmailAddressPolicy | Sort-Object Priority

  $ObjectCount = ($EmailAddressPolicies | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Policy in $EmailAddressPolicies) {
  
    $HashTable += @{ 
      Name = $Policy.Name; 
      Priority = $Policy.Priority; 
      RecipientFilter = $Policy.RecipientFilter; 
      RecipientFilterApplied = $Policy.RecipientFilterApplied;
    }
  }

  $SectionTitle = 'Email Address Policy'

  $PolicyText = 'policies'
  if($ObjectCount-eq 1) {$PolicyText = 'policy'}

  $Text = ('The Exchange Organization contains {0} email address {1}.' -f $ObjectCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Priority, RecipientFilter, RecipientFilterApplied `
    -Headers 'Name', 'Priority', 'Recipient Filter', 'Recipient Filter Applied' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 140
    $Table.Columns.Item(2).Width = 50
    $Table.Columns.Item(3).Width = 200
    $Table.Columns.Item(4).Width = 50

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
    
    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Email Address Policy Details'

      foreach ($Policy in $EmailAddressPolicies) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Policy.Name)

        $ObjectInformation = Expand-Object -Object $Policy

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}

function Get-ReceiveConnectorInformation {

  Show-ProgressBar -Status 'Mail Flow - Receive Connectors' -PercentComplete 15 -Stage 1

  $ReceiveConnectors = Get-ReceiveConnector | Sort-Object Server,Name

  $ObjectCount = ($ReceiveConnectors | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Connector in $ReceiveConnectors) {
  
    $HashTable += @{ 
      Server = $Connector.Server; 
      Name = $Connector.Name; 
      Bindings = ($Connector.Bindings -join ', '); 
      Enabled = $Connector.Enabled; 
    }
  }

  $SectionTitle = 'Receive Connector'

  $PolicyText = 'connectors'
  if($ObjectCount-eq 1) {$PolicyText = 'connector'}

  $Text = ('The Exchange Organization contains {0} receive {1} .' -f $ObjectCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Server, Name, Bindings, Enabled `
    -Headers 'Server', 'Name', 'Bindings', 'Enabled' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 80
    $Table.Columns.Item(2).Width = 150
    $Table.Columns.Item(3).Width = 150
    $Table.Columns.Item(4).Width = 50

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Receive Connector Details'

      foreach ($Connector in $ReceiveConnectors) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Connector.Identity)

        $ObjectInformation = Expand-Object -Object $Connector 

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}

function Get-SendConnectorInformation {

  Show-ProgressBar -Status 'Mail Flow - Send Connectors' -PercentComplete 15 -Stage 1

  $SendConnectors = Get-SendConnector | Sort-Object Name

  $ObjectCount = ($SendConnectors | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Connector in $SendConnectors) {
  
    $HashTable += @{ 
      Name = $Connector.Name; 
      AddressSpaces = ($Connector.AddressSpaces -join ', '); 
      SourceTransportServers = ($Connector.SourceTransportServers -join ', ');
      Enabled = $Connector.Enabled; 
    }
  }

  $SectionTitle = 'Send Connector'

  $PolicyText = 'connectors'
  if($ObjectCount-eq 1) {$PolicyText = 'connector'}

  $Text = ('The Exchange Organization contains {0} send {1} .' -f $ObjectCount, $PolicyText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, AddressSpaces, SourceTransportServers, Enabled `
    -Headers 'Name', 'Address Spaces', 'Source Transport Servers', 'Enabled' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 120
    $Table.Columns.Item(2).Width = 150
    $Table.Columns.Item(3).Width = 150
    $Table.Columns.Item(4).Width = 50

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) { 

      Write-WordLine -Style 3 -Tabs 0 -Name 'Send Connector Details'

      foreach ($Connector in $SendConnectors) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $Connector.Identity)

        $ObjectInformation = Expand-Object -Object $Connector 

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}

function Get-PublicFolderInformation {

  if($IncludePublicFolders) {

    Show-ProgressBar -Status 'Public Folders - Public Folder Details (long running)' -PercentComplete 15 -Stage 1

    # This code will run forvever so it's not enabled by default
    # Get-PublicFolderStatistics -ResultSize Unlimited | Select-Object Name,DatabaseName,TotalItemSize | Measure-Object -Property TotalItemSize -Sum

    $PublicFolderCount = ((Get-PublicFolder "\" -Recurse -ResultSize Unlimited) | Measure-Object).Count

    $MailPublicFolders = Get-MailPublicFolder -ResultSize Unlimited
    $MailPublicFolderCount = ($MailPublicFolders | Measure-Object).Count

    $RootPublicFolders = Get-PublicFolder "\" -GetChildren
    $RootPublicFolderCount = ($RootPublicFolders | Measure-Object).Count

    [Collections.Hashtable[]]$PublicFolderInformation = @()
  
    $PublicFolderInformation += @{ Data = "Public Folders in Root"; Value = $RootPublicFolderCount; }
    $PublicFolderInformation += @{ Data = "Public Folders"; Value = $PublicFolderCount; }
    $PublicFolderInformation += @{ Data = "Mail enabled Public Folders"; Value = $MailPublicFolderCount; }

    $Table = Add-WordTable -Hashtable $PublicFolderInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

    Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 300
    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
    
    if($RootPublicFolderCount -gt 0) { 

      Write-WordLine -Style 2 -Tabs 0 -Name 'Root Public Folders'

      [Collections.Hashtable[]]$HashTable = @()

      foreach($PublicFolder in $RootPublicFolders) {

        $HashTable += @{
          Name = $PublicFolder.Name;
          HasSubFolders = $PublicFolder.HasSubFolders;
          ContentMailboxName = $PublicFolder.ContentMailboxName;
        }
      }

      $Table = Add-WordTable -Hashtable $HashTable `
      -Columns Name, HasSubFolders, ContentMailboxName `
      -Headers 'Name', 'Has SubFolders', 'Content MailboxName' `
      -AutoFit $wdAutoFitFixed

      Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

      $Table.Columns.Item(1).Width = 150
      $Table.Columns.Item(2).Width = 80
      $Table.Columns.Item(3).Width = 150

      $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

      Select-WordEndOfDocument

      $Table = $Null

      Write-EmptyWordLine
    }
    else {
      Write-WordLine -Style 0 -Tabs 0 -Name "The public folder hierarchy is empty."
    }

    if($MailPublicFolderCount -gt 0) {

      Write-WordLine -Style 2 -Tabs 0 -Name 'Mail Public Folders'

      [Collections.Hashtable[]]$HashTable = @()

      foreach($PublicFolder in $MailPublicFolders) {

        $HashTable += @{
          DisplayName = $PublicFolder.DisplayName;
          PrimarySmtpAddress = $PublicFolder.PrimarySmtpAddress;
          EmailAddresses = $PublicFolder.EmailAddresses -join ', ';
        }
      }

      $Table = Add-WordTable -Hashtable $HashTable `
      -Columns DisplayName, PrimarySmtpAddress, EmailAddresses `
      -Headers 'Display Name', 'Primary SmtpAddress', 'Email Addresses' `
      -AutoFit $wdAutoFitFixed

      Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

      Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Size $WordSmallFontSize
      Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize
      Set-WordCellFormat -Collection $Table.Columns.Item(3).Cells -Size $WordSmallFontSize

      $Table.Columns.Item(1).Width = 120
      $Table.Columns.Item(2).Width = 120
      $Table.Columns.Item(3).Width = 190

      $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

      Select-WordEndOfDocument

      $Table = $Null

      Write-EmptyWordLine
      
    }
    else {
      Write-WordLine -Style 0 -Tabs 0 -Name 'The public folder hierarchy does contain any mail enabled public folders.'
    }
  }
  else {
    $Text = 'This document does not include detailed information about the public folder hierarchy. To include detailed information run the script using the -IncludePublicFolders switch.'

    Write-WordLine -Style 0 -Tabs 0 -Name $Text
  }
}

function Get-PublicFolderMailboxInformation {

  Show-ProgressBar -Status 'Public Folder - Public Folder Mailboxes' -PercentComplete 15 -Stage 1

  $PublicFolderMailboxes = Get-Mailbox -PublicFolder -ResultSize Unlimited | Sort-Object

  $ObjectCount = ($PublicFolderMailboxes | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Mailbox in $PublicFolderMailboxes) {
  
    $HashTable += @{ 
      Name = $Mailbox.Name; 
      IsRootPublicFolderMailbox = $Mailbox.IsRootPublicFolderMailbox;
      IsExcludedFromServingHierarchy = $Mailbox.IsExcludedFromServingHierarchy;
      Database = $Mailbox.Database; 
    }
  }

  $SectionTitle = 'Public Folder Mailboxes'

  if($ObjectCount -eq 0) {
    $Text = 'The Exchange Organization has no public folder mailboxes.'
  }
  else { 

    $ObjectText = 'mailboxes'
    if($ObjectCount-eq 1) {$ObjectText = 'mailbox'}

    $Text = ('The Exchange Organization contains {0} public folder {1} .' -f $ObjectCount, $ObjectText)
  }

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, IsRootPublicFolderMailbox, IsExcludedFromServingHierarchy, Database `
    -Headers 'Name', 'Primary Hierarchy', 'Excluded From Serving Hierarchy', 'Database' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 100
    $Table.Columns.Item(2).Width = 80
    $Table.Columns.Item(3).Width = 80
    $Table.Columns.Item(4).Width = 100

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

function Get-UMDialPlanInformation {

  Show-ProgressBar -Status 'Unified Messaging - Dial Plans' -PercentComplete 15 -Stage 1

  $DialPlans = Get-UMDialPlan

  $ObjectCount = ($DialPlans | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($DialPlan in $DialPlans) {
  
    $HashTable += @{ 
      Name = $DialPlan.Name; 
      AddressSpaces = ($DialPlan.AddressSpaces -join ', '); 
      SourceTransportServers = ($DialPlan.SourceTransportServers -join ', ');
      Enabled = $DialPlan.Enabled; 
    }
  }

  $SectionTitle = 'Dial Plans'

  if($ObjectCount -eq 0) {
    $text = 'The Exchange Organization contains no dial plans.'
  }
  else {
    $PolicyText = 'plans'
    if($ObjectCount-eq 1) {$PolicyText = 'plan'}

    $Text = ('The Exchange Organization contains {0} dial {1} .' -f $ObjectCount, $PolicyText)

  }


  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    if($ObjectCount -ne 0) { 

      <#
    
          $Table = Add-WordTable -Hashtable $HashTable `
          -Columns Name, AddressSpaces, SourceTransportServers, Enabled `
          -Headers 'Name', 'Address Spaces', 'Source Transport Servers', 'Enabled' `
          -AutoFit $wdAutoFitFixed

          Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

          $Table.Columns.Item(1).Width = 120
          $Table.Columns.Item(2).Width = 150
          $Table.Columns.Item(3).Width = 150
          $Table.Columns.Item(4).Width = 50

          $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

          Select-WordEndOfDocument

          $Table = $Null

          Write-EmptyWordLine

      #>
    }
  }
}

function Get-DiskSpace {
  param(
    [string]$ServerName
  )
    
  $wmiObject = Get-WmiObject -Class Win32_Volume -ComputerName $ServerName | Select-Object Name, Capacity, FreeSpace, BootVolume, SystemVolume, FileSystem | Sort-Object -Property Name 

  [Collections.Hashtable[]]$HashTable = @()

  foreach($Drive in $wmiObject) {

    $HashTable += @{
      Name =$Drive.Name;
      Capacity = [decimal]::round($Drive.Capacity/1GB);
      FreeDiskSpace = [decimal]::round($Drive.FreeSpace/1GB);
      BootVolume = $Drive.BootVolume;
      SystemVolume = $Drive.SystemVolume;
      FileSystem = $Drive.FileSystem;
    }
  }

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 4 -Tabs 0 -Name 'Disk Volumes'

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Capacity, FreeDiskSpace, BootVolume, SystemVolume,FileSystem `
    -Headers 'Name', 'Capacity (GB)', 'Free (GB)', 'Boot Volume', 'System Volume', 'File System' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 130
    $Table.Columns.Item(2).Width = 70
    $Table.Columns.Item(3).Width = 70
    $Table.Columns.Item(4).Width = 60
    $Table.Columns.Item(5).Width = 60
    $Table.Columns.Item(6).Width = 60

    # set font to 8pt
    1..6 | %{ Set-WordCellFormat -Collection $Table.Columns.Item($_).Cells -Size $WordSmallFontSize }

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null
  }
}

Function Get-OperatingSystemVersion {
  param(
    [Parameter(Mandatory=$true)][string]$OSVersion
  )
    
    switch($OSVersion) {
        "6.0.6000" {return [OrgReport.OSVersionName]::Windows2008}
        "6.1.7600" {return [OrgReport.OSVersionName]::Windows2008R2}
        "6.1.7601" {return [OrgReport.OSVersionName]::Windows2008R2}
        "6.2.9200" {return [OrgReport.OSVersionName]::Windows2012}
        "6.3.9600" {return [OrgReport.OSVersionName]::Windows2012R2}
        "10.0.14393" {return [OrgReport.OSVersionName]::Windows2016}
        "10.0.17713" {return [OrgReport.OSVersionName]::Windows2019}
        default{return [OrgReport.OSVersionName]::Unknown}
    }
}



function Get-NetworkInformation {
  param(
    [Parameter(Mandatory)][string]$ServerName
  )

}

function Get-OperatingSystemInformation {
  param(
    [Parameter(Mandatory)][string]$ServerName
  )

  [OrgReport.OperatingSystemObject]$OperatingSystem = New-Object -TypeName OrgReport.OperatingSystemObject
  $WmiOS = Get-WmiObject -ComputerName $ServerName -Class Win32_OperatingSystem

  $OperatingSystem.OSVersionBuild = $WmiOS.Version
  $OperatingSystem.OSVersion = Get-OperatingSystemVersion -OSVersion $OperatingSystem.OSVersionBuild
  $OperatingSystem.OSName = $WmiOS.Caption
  $OperatingSystem.OperatingSystem = $OperatingSystem

  $CurrentDateTime = Get-Date
  $ServerUptime = [Management.ManagementDateTimeConverter]::ToDateTime($WmiOS.LastBootUpTime)

  $OperatingSystem.BootUpTimeInDays = ($CurrentDateTime - $ServerUptime).Days
  $OperatingSystem.BootUpTimeInHours = ($CurrentDateTime - $ServerUptime).Hours
  $OperatingSystem.BootUpTimeInMinutes = ($CurrentDateTime - $ServerUptime).Minutes
  $OperatingSystem.BootUpTimeInSeconds = ($CurrentDateTime - $ServerUptime).Seconds

  if($env:COMPUTERNAME -eq $ServerName) {
    $OperatingSystem.TimeZone = ([System.TimeZoneInfo]::Local).StandardName
  }
  else {
    $OperatingSystem.TimeZone = Invoke-Command -ComputerName $ServerName -ScriptBlock {([System.TimeZoneInfo]::Local).StandardName}
  }

  return $OperatingSystem
}

function Write-ServerDetails {
  param(
    [Parameter(Mandatory)][OrgReport.ExchangeServerObject]$ServerObject
  )

  [Collections.Hashtable[]]$ObjectInformation = @()

  $ObjectInformation += @{ Data = 'Operating System'; Value = $ServerObject.OperatingSystem.OSName; }
  $ObjectInformation += @{ Data = 'Operating Version'; Value = $ServerObject.OperatingSystem.OSVersionBuild; }
  $ObjectInformation += @{ Data = 'Up Time'; Value = ("{0} day(s), {1} hour(s), {2} minute(s), {3} second(s)" -f $ServerObject.OperatingSystem.BootUpTimeInDays,$ServerObject.OperatingSystem.BootUpTimeInHours, $ServerObject.OperatingSystem.BootUpTimeInMinutes, $ServerObject.OperatingSystem.BootUpTimeInSeconds ) }
  $ObjectInformation += @{ Data = 'Time Zone'; Value = $ServerObject.OperatingSystem.TimeZone; }

  if($ExportTo -eq 'MSWord') { 

    $Table = Add-WordTable -Hashtable $ObjectInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed

		Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

		$Table.Columns.Item(1).Width = 175
		$Table.Columns.Item(2).Width = 300

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		Select-WordEndOfDocument

		$Table = $Null

    Write-EmptyWordLine
  }
}

function Get-ServerDetails {
  param(
    [Parameter(Mandatory)][Microsoft.Exchange.Data.Directory.Management.ExchangeServer]$ExchangeServer,
    [switch]$IsEdgeServer
  )

  [OrgReport.ExchangeServerObject]$ServerObject = New-Object -TypeName OrgReport.ExchangeServerObject
  $ServerObject.Name = $ExchangeServer.Name
  $ServerObject.OperatingSystem = Get-OperatingSystemInformation -ServerName $ExchangeServer.Name

  Write-ServerDetails -ServerObject $ServerObject

  Get-DiskSpace -ServerName $ExchangeServer.Name
}


function Get-ExchangeServerInformation {

  Show-ProgressBar -Status 'Servers - Servers' -PercentComplete 15 -Stage 1

  $ExchangeServers = Get-ExchangeServer | Sort-Object Name

  $ObjectCount = ($ExchangeServers | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Server in $ExchangeServers) {
  
    $HashTable += @{ 
      Name = $Server.Name; 
      ServerRole = $Server.ServerRole; 
      Edition = $Server.Edition;
      Version = ('{0}.{1}.{2}' -f $Server.AdminDisplayVersion.Major, $Server.AdminDisplayVersion.Minor, $Server.AdminDisplayVersion.Build)
    }
  }

  $SectionTitle = 'Exchange Servers'

  $ObjectText = 'servers'

  if($ObjectCount-eq 1) {$ObjectText = 'server'}

  $Text = ('The Exchange Organization contains {0} Exchange {1} .' -f $ObjectCount, $ObjectText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, ServerRole, Edition `
    -Headers 'Name', 'Server Role', 'Edition' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 100
    $Table.Columns.Item(2).Width = 80
    $Table.Columns.Item(3).Width = 80

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    # process each server
    foreach($Server in $ExchangeServers) {

      Write-WordLine -Style 3 -Tabs 0 -Name $Server.Name

      if($Server.ServerRole -contains 'Edge') {
        Get-ServerDetails -ExchangeServer $Server -IsEdgeServer
      }
      else {
        Get-ServerDetails -ExchangeServer $Server
      }
      
      Write-EmptyWordLine

    }
  }

}

function Get-DatabaseInformation {

  Show-ProgressBar -Status 'Servers - Databases' -PercentComplete 15 -Stage 1

  $Databases = Get-MailboxDatabase -Status | Sort-Object -Property Name

  $ObjectCount = ($Databases | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()
  
  $CultureInfo = ([System.Globalization.CultureInfo]::GetCultureInfo(1031))

  foreach ($Database in $Databases) {
  
    $HashTable += @{ 
      Name = $Database.Name;
      ActiveServer = $Database.Server.Name;
      Mounted = $Database.Mounted;
      Copies = ($Database.AllDatabaseCopies | Select HOstServerName | Sort-Object).HostServerName -join ', '
      Size = ($Database.DatabaseSize.ToBytes() / 1GB).ToString('N0', $CultureInfo)
    }
  }

  $SectionTitle = 'Databases'

  $ObjectText = 'databases'

  if($ObjectCount-eq 1) {$ObjectText = 'database'}

  $Text = ('The Exchange Organization contains {0} mailbox {1} .' -f $ObjectCount, $ObjectText)

  if($ExportTo -eq 'MSWord') {

    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine
    
    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, ActiveServer, Mounted, Size, Copies `
    -Headers 'Name', 'Active Server', 'Mounted', 'Size (GB)', 'Database Copies' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 80
    $Table.Columns.Item(2).Width = 80
    $Table.Columns.Item(3).Width = 80
    $Table.Columns.Item(4).Width = 80
    $Table.Columns.Item(5).Width = 150

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

  }
}

function Get-DatabaseAvailabilityGroupInformation {
  Show-ProgressBar -Status 'Servers - Database Availability Groups' -PercentComplete 15 -Stage 1

  $DAGs= Get-DatabaseAvailabilityGroup -Status | Sort-Object Name
  $ObjectCount = ($DAGs | Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach($DAG in $DAGs) {
     $HashTable += @{ 
      Name = $DAG.Name;
      Servers = (($DAG.Servers | Select Name | Sort-Object -Property Name).Name) -join ', ';
    }
  }  

  $SectionTitle = 'Database Availability Groups'

  $ObjectText = 'groups'
  if($ObjectCount-eq 1) {$ObjectText = 'group'}
  $Text = ('The Exchange Organization contains {0} database availability {1}.' -f $ObjectCount, $ObjectText)

  if($ExportTo -eq 'MSWord') {
    Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 0 -Tabs 0 -Name $Text

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Servers `
    -Headers 'Name', 'Member Servers' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 100
    $Table.Columns.Item(2).Width = 150

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine

    if($IncludeDetails) {

      Write-WordLine -Style 3 -Tabs 0 -Name 'DAG Details'

      foreach ($DAG in $DAGs) {

        Write-WordLine -Style 0 -Tabs 0 -Name ('Policy: {0}' -f $DAG.Identity)

        $ObjectInformation = Expand-Object -Object $DAG 

        $Table = Add-WordTable -Hashtable $ObjectInformation -Columns Data,Value -List -Format $wdTableGrid -AutoFit $wdAutoFitFixed 

        Set-WordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15 -Size $WordSmallFontSize
        Set-WordCellFormat -Collection $Table.Columns.Item(2).Cells -Size $WordSmallFontSize

        $Table.Columns.Item(1).Width = 180
        $Table.Columns.Item(2).Width = 300
        $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

        Select-WordEndOfDocument

        $Table = $Null

        Write-EmptyWordLine 
     
      }
    }
  }
}

function Get-VirtualDirectoryInformation {
  param(
    [string]$Source,
    [string]$Sort,
    [string]$Title,
    [string]$Command
  )

  Show-ProgressBar -Status 'Servers - Virtual Directory (AutoDiscover)' -PercentComplete 15 -Stage 1
 
  $expr = ('Get-{0} | Sort-Object {1}' -f $Source, $Sort)

  $VirtualDirectories = Invoke-Expression $expr
  $ObjectCount = ($VirtualDirectories| Measure-Object).Count

  [Collections.Hashtable[]]$HashTable = @()

  foreach ($Dir in $VirtualDirectories) {
      $HashTable += @{ 
      Name = $Dir.Name;
      Server = $Dir.Server;
      InternalUrl = $Dir.InternalUrl;
      ExternalUrl = $DIr.ExternalUrl;
    }
  }

  if($ExportTo -eq 'MSWord') {
    
    Write-WordLine -Style 3 -Tabs 0 -Name $Title

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $HashTable `
    -Columns Name, Server, InternalUrl, ExternalUrl `
    -Headers 'Name', 'Server', 'Internal Url', 'External Url' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15

    $Table.Columns.Item(1).Width = 100
    $Table.Columns.Item(2).Width = 80
    $Table.Columns.Item(3).Width = 150
    $Table.Columns.Item(4).Width = 150
    
    1..4 | %{ Set-WordCellFormat -Collection $Table.Columns.Item($_).Cells -Size $WordSmallFontSize}

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

function Get-ExchangeVirtualDirectories {

  $SectionTitle = 'Virtual Directories'
  Write-WordLine -Style 2 -Tabs 0 -Name $SectionTitle
  Write-EmptyWordLine

  $infoSources = [ordered]@{
    'AutodiscoverVirtualDirectory' = @{Sort = 'Server';Title = 'AutoDiscover Virtual Directory';Command = '';}
    'OwaVirtualDirectory' = @{Sort = 'Server';Title = 'OWA Virtual Directory';Command = '';}
    'EcpVirtualDirectory' = @{Sort = 'Server';Title = 'ECP Virtual Directory';Command = '';}
    'MapiVirtualDirectory' = @{Sort = 'Server';Title = 'MAPI Virtual Directory';Command = '?{$_.AdminDisplayVersion -ilike ''*15''}';}
    'WebServicesVirtualDirectory' = @{Sort = 'Server';Title = 'ECP Virtual Directory';Command = '';}
    'OabVirtualDirectory' = @{Sort = 'Server';Title = 'ECP Virtual Directory';Command = '';}
    'PowerShellVirtualDirectory' = @{Sort = 'Server';Title = 'ECP Virtual Directory';Command = '';}
  }

  $infoSources.GetEnumerator() | ForEach-Object { 
    Get-VirtualDirectoryInformation -Source $_.Key -Sort $infoSources[$_.Key].Sort -Title $infoSources[$_.Key].Title -Command infoSources[$_.Key].Command
  } 

}

function Output-IntroductionText {

  if($IncludeIntroduction) {

    Write-WordLine -Style 1 -Tabs 0 -Name 'Introduction'

    $Text = 'Some introductory text...'

    Write-WordLine -Style 0 -Tabs 0 -Name $Text 
    
  }

}

function Get-WordDocumentationLinks {

  # todo add proper hyperlinks as table cell information to Word

  [Collections.Hashtable[]]$HashTableDocumentationLinks = @()

  $HashTableDocumentationLinks += @{Title='Link 1';Link = 'http://www.google.de'}

  $SectionTitle = 'Appendix'

  if($ExportTo -eq 'MSWord') {

    $Script:Selection.InsertNewPage()
    
    Write-WordLine -Style 1 -Tabs 0 -Name $SectionTitle

    Write-WordLine -Style 3 -Tabs 0 -Name 'Documentation Links'

    Write-EmptyWordLine

    $Table = Add-WordTable -Hashtable $HashTableDocumentationLinks `
    -Columns Title, Link `
    -Headers 'Title','Link' `
    -AutoFit $wdAutoFitFixed

    Set-WordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

    $Table.Columns.Item(1).Width = 180
    $Table.Columns.Item(2).Width = 200

    $Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

    Select-WordEndOfDocument

    $Table = $Null

    Write-EmptyWordLine
  }
}

#endregion


#region ideas

# placeholder region for further optimization of the script

#endregion

### MAIN #########################

$script:StartTime = Get-Date

# Let's do some intital checking first

[string]$Script:CurrentOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

$Script:CoName = $CompanyName
Write-Verbose -Message "$(Get-Date): Company name (CoName) is $($Script:CoName)"

#region Word variables

if($ExportTo -eq 'MSWord') {

  # Test, if we can write to Word
  # Test-WordPrerequisites
  
  # default values
  [int]$WordDefaultFontSize = 11
  [int]$WordSmallFontSize = 8

  # Word Enumerated Constants
  # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa211923(v=office.11)
  [int]$wdAlignPageNumberRight = 2
  [long]$wdColorGray15 = 14277081
  [long]$wdColorGray05 = 15987699 
  [int]$wdMove = 0
  [int]$wdSeekMainDocument = 0
  [int]$wdSeekPrimaryFooter = 4
  [int]$wdStory = 6
  [long]$wdColorRed = 255
  [int]$wdColorBlack = 0
  [long]$wdColorYellow = 65535 
  [int]$wdWord2007 = 12
  [int]$wdWord2010 = 14
  [int]$wdWord2013 = 15
  [int]$wdWord2016 = 16
  [int]$wdFormatDocumentDefault = 16
  [int]$wdFormatPDF = 17
  
  # Word Paragraph Alignment
  # https://devblogs.microsoft.com/scripting/how-can-i-right-align-a-single-column-in-a-word-table/
  # https://docs.microsoft.com/en-us/office/vba/api/Word.WdParagraphAlignment
  [int]$wdAlignParagraphLeft = 0
  [int]$wdAlignParagraphCenter = 1
  [int]$wdAlignParagraphRight = 2

  # Word Cell Certical Alignment
  # https://docs.microsoft.com/en-us/office/vba/api/Word.WdCellVerticalAlignment
  [int]$wdCellAlignVerticalTop = 0
  [int]$wdCellAlignVerticalCenter = 1
  [int]$wdCellAlignVerticalBottom = 2

  # Word AutoFit Behavior
  # https://docs.microsoft.com/en-us/office/vba/api/Word.WdAutoFitBehavior
  [int]$wdAutoFitFixed = 0
  [int]$wdAutoFitContent = 1
  [int]$wdAutoFitWindow = 2

  # Word RulerStyle
  # https://docs.microsoft.com/en-us/office/vba/api/Word.WdRulerStyle
  [int]$wdAdjustNone = 0
  [int]$wdAdjustProportional = 1
  [int]$wdAdjustFirstColumn = 2
  [int]$wdAdjustSameWidth = 3

  [int]$PointsPerTabStop = 36
  [int]$Indent0TabStops = 0 * $PointsPerTabStop
  [int]$Indent1TabStops = 1 * $PointsPerTabStop
  [int]$Indent2TabStops = 2 * $PointsPerTabStop
  [int]$Indent3TabStops = 3 * $PointsPerTabStop
  [int]$Indent4TabStops = 4 * $PointsPerTabStop

  # Word Style Names in English, Danish, German, French 
  # http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
  [int]$wdStyleHeading1 = -2
  [int]$wdStyleHeading2 = -3
  [int]$wdStyleHeading3 = -4
  [int]$wdStyleHeading4 = -5
  [int]$wdStyleNoSpacing = -158
  [int]$wdTableGrid = -155

  # URL non existent
  # http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
  [int]$wdLineStyleNone = 0
  [int]$wdLineStyleSingle = 1

  [int]$wdHeadingFormatTrue = -1
  [int]$wdHeadingFormatFalse = 0 
}

#endregion

#region Html variable

if($ExportTo -eq 'Html') {
  Set-Variable HtmlRedMask -Option AllScope -Value "#FF0000" 4>$Null
  Set-Variable HtmlCyanMask -Option AllScope -Value "#00FFFF" 4>$Null
  Set-Variable HtmlBlueMask -Option AllScope -Value "#0000FF" 4>$Null
  Set-Variable HtmlDarkBlueMask -Option AllScope -Value "#0000A0" 4>$Null
  Set-Variable HtmlLightBlueMask -Option AllScope -Value "#ADD8E6" 4>$Null
  Set-Variable HtmlPurpleMask -Option AllScope -Value "#800080" 4>$Null
  Set-Variable HtmlYellowMask -Option AllScope -Value "#FFFF00" 4>$Null
  Set-Variable HtmlLimeMask -Option AllScope -Value "#00FF00" 4>$Null
  Set-Variable HtmlMagentaMask -Option AllScope -Value "#FF00FF" 4>$Null
  Set-Variable HtmlWhiteMask -Option AllScope -Value "#FFFFFF" 4>$Null
  Set-Variable HtmlSilverMask -Option AllScope -Value "#C0C0C0" 4>$Null
  Set-Variable HtmlGrayMask -Option AllScope -Value "#808080" 4>$Null
  Set-Variable HtmlBlackMask -Option AllScope -Value "#000000" 4>$Null
  Set-Variable HtmlOrangeMask -Option AllScope -Value "#FFA500" 4>$Null
  Set-Variable HtmlMaroonMask -Option AllScope -Value "#800000" 4>$Null
  Set-Variable HtmlGreenMask -Option AllScope -Value "#008000" 4>$Null
  Set-Variable HtmlOliveMask -Option AllScope -Value "#808000" 4>$Null

  Set-Variable HtmlBold -Option AllScope -Value 1 4>$Null
  Set-Variable HtmlItalics -Option AllScope -Value 2 4>$Null
  Set-Variable HtmlRed -Option AllScope -Value 4 4>$Null
  Set-Variable HtmlCyan -Option AllScope -Value 8 4>$Null
  Set-Variable HtmlBlue -Option AllScope -Value 16 4>$Null
  Set-Variable HtmlDarkBlue -Option AllScope -Value 32 4>$Null
  Set-Variable HtmlLighBblue -Option AllScope -Value 64 4>$Null
  Set-Variable HtmlPurple -Option AllScope -Value 128 4>$Null
  Set-Variable HtmlYellow -Option AllScope -Value 256 4>$Null
  Set-Variable HtmlLime -Option AllScope -Value 512 4>$Null
  Set-Variable HtmlMagenta -Option AllScope -Value 1024 4>$Null
  Set-Variable HtmlWhite -Option AllScope -Value 2048 4>$Null
  Set-Variable HtmlSilver -Option AllScope -Value 4096 4>$Null
  Set-Variable HtmlGray -Option AllScope -Value 8192 4>$Null
  Set-Variable HtmlOlive -Option AllScope -Value 16384 4>$Null
  Set-Variable HtmlOrange -Option AllScope -Value 32768 4>$Null
  Set-Variable HtmlMaroon -Option AllScope -Value 65536 4>$Null
  Set-Variable HtmlGreen -Option AllScope -Value 131072 4>$Null
  Set-Variable HtmlBlack -Option AllScope -Value 262144 4>$Null
}

#endregion

# Let's begin with the real stuff ###################################

$DocumentTitle = 'Microsoft Exchange Organization Report'
$AbstractTitle = 'Microsoft Exchange Organization Report'
$SubjectTitle = 'Active Directory Inventory Report'
$UserName = $env:username

if($ADForest -ne ''-and $ADDomain -ne '') {
  $ADForest = $ADDomain
}

# Exchange related variables
[bool]$E2010 = $true
[bool]$E2013 = $false

Test-ExchangeManagementShellVersion

if ($E2010) {
  Set-ADServerSettings -ViewEntireForest:$ViewEntireForest
} 
else {
  $global:AdminSessionADSettings.ViewEntireForest = $ViewEntireForest
}

Set-ADServerSettings -ViewEntireForest $true

if($IncludePublicFolders) {
  Write-Warning -Message 'You''ve selected to include public folders. This will effect the overall runtime of the script, depending on the size of the public folder hierarchy!'
}

switch ($ExportTo) {
  'MSWord' {

    # Prepare a new Word document for our report
    New-MicrosoftWordDocument
  
    Output-IntroductionText

    # Fetch general Active Directory Information first
    Get-ActiveDirectoryInformation

    # Let's work on Exchange Org information

    Set-SectionTitle -Style 1 -SectionTitle 'Exchange Organization Overview' -NewPage

    Get-ExchangeOrganizationConfig

    Get-RecipientInformation

    Get-AcceptedDomainInformation

    Get-DatabaseOverviewInformation

    # Get-TransportConfigInformation

    # Permissions

    Set-SectionTitle -Style 1 -SectionTitle 'Permissions' -NewPage

    Get-AdminPermissionInformation

    Get-UserRoleAssignmentPolicies

    Get-OutlookWebAppPolicies

    # Compliance Management

    Set-SectionTitle -Style 1 -SectionTitle 'Compliance Management' -NewPage

    Get-ComplianceInformation

    Get-RetentionPolicyInformation

    Get-RetentionPolicyTagInformation

    Get-JournalingInformation

    # Organization

    Set-SectionTitle -Style 1 -SectionTitle 'Organization' -NewPage

    Get-SharingInformation

    Get-AppInformation

    Get-AddressListInformation

    # Protection

    Set-SectionTitle -Style 1 -SectionTitle 'Protection' -NewPage

    Get-MalwareInformation
    
    ## Malware Rules 

    # Mail Flow

    Set-SectionTitle -Style 1 -SectionTitle 'Mail Flow' -NewPage

    Get-TransportRuleInformation
    
    ## Accepted Domains (already covered)his step)

    Get-EmailAddressPolicyInformation

    Get-ReceiveConnectorInformation

    Get-SendConnectorInformation


    # Mobile Devices

    Set-SectionTitle -Style 1 -SectionTitle 'Mobile' -NewPage

    Get-MobileDeviceInformation

    Get-MobileDevicePolicies

    # Public Folders 

    Set-SectionTitle -Style 1 -SectionTitle 'Public Folder' -NewPage

    Get-PublicFolderInformation

    Get-PublicFolderMailboxInformation

    
    # Unified Messaging

    Set-SectionTitle -Style 1 -SectionTitle 'Unified Messaging' -NewPage

    Get-UMDialPlanInformation

    ## UM IP Gateways

    # Servers

    Set-SectionTitle -Style 1 -SectionTitle 'Servers' -NewPage

    Get-ExchangeServerInformation

    Get-DatabaseInformation

    Get-DatabaseAvailabilityGroupInformation

    Get-ExchangeVirtualDirectories

    ## Certificates

    # Global Overrides

    # Hybrid

    # Finalize Word document
    Get-WordDocumentationLinks

    Update-DocumentProperty -DocumentTitle $DocumentTitle -AbstractTitle $AbstractTitle -SubjectTitle $SubjectTitle -Author $UserName 

    # That's it, close the document
    Close-WordDocument

  }
  'Html' {
    # To-Do: Enhanced Html Output
  }

}

$StopWatch.Stop()

Write-Output ('The script runtime was {0}min {1}s' -f $StopWatch.Elapsed.Minutes, $StopWatch.Elapsed.Seconds)