<#PSScriptInfo

.VERSION 2.5.0.3

.GUID b4e8f9aa-1433-4d8b-8aea-8681fbdfde8c

.AUTHOR Microsoft Corporation

.COMPANYNAME Microsoft Corporation

.EXTERNALMODULEDEPENDENCIES

.TAGS SharePoint,ReverseDSC,DesiredStateConfiguration,DSC,DSCResourceKit,DSCResource

.ICONURI https://GitHub.com/Microsoft/SharePointDSC.Reverse/blob/master/Images/SharePointDSC.Reverse.png?raw=true

.RELEASENOTES

* Updated Requirement for SharePointDSC 2.5.0.0;
* Fixes issue with Central Admin Invalid Port;
* Now filters out invalid usernames (e.g. true) from Web App Policies;
* Added support for the SPSiteUrl Resource;
* Fixed and issue with retrieving Search Service Application without the ApplicationPool specified;
* Fixed an issue with multiple lines in SPSite and SPWeb descriptions;
* Fixed issue with Other languages than english when extracting configuration database;
* Fixed issue where an invalid DSC block structure was sent for SPStateServiceApp;
* Changed behavior for Site Collection Owner. If Service account, use variable, otherwise plaintext;
#>

#Requires -Modules @{ModuleName="ReverseDSC";ModuleVersion="1.9.2.9"},@{ModuleName="SharePointDSC";ModuleVersion="2.5.0.0"}

<#

.DESCRIPTION
 Extracts the DSC Configuration of an existing SharePoint environment, allowing you to analyze it or to replicate the farm.

#>

param(
    [ValidateSet("Lite","Default", "Full")]
    [System.String]$Mode = "Default",
    [switch]$Standalone,
    [Boolean]$Confirm = $true,
    [String]$OutputFile = $null,
    [switch]$SkipSitesAndWebs = $false,
    [switch]$Azure = $false)

<## Script Settings #>
$VerbosePreference = "SilentlyContinue"

<## Dependency Hashes ##>
$Script:DH_SPQUOTATEMPLATE = @{}

<## Scripts Variables #>
$Script:dscConfigContent = ""
$Script:ErrorLog = ""
$Script:configName = ""
$Script:currentServerName = ""
$SPDSCSource = "$env:ProgramFiles\WindowsPowerShell\Modules\SharePointDSC\"
$SPDSCVersion = "2.5.0.0"
$Script:spCentralAdmin = ""
$Script:ExtractionModeValue = "2"
$script:SkipSitesAndWebs = $SkipSitesAndWebs
if($Mode.ToLower() -eq "lite")
{
    $Script:ExtractionModeValue = 1
}
elseif($Mode.ToLower() -eq "full")
{
    $Script:ExtractionModeValue = 3
}

try
{
    $currentScript = Test-ScriptFileInfo $SCRIPT:MyInvocation.MyCommand.Path
    $Script:version = $currentScript.Version.ToString()
}
catch
{
    $Script:version = "N/A"
}
$Script:SPDSCPath = $SPDSCSource + $SPDSCVersion
$Global:spFarmAccount = ""


<## This is the main function for this script. It acts as a call dispatcher, calling the various functions required in the proper order to get the full farm picture. #>
function Orchestrator
{
    Test-Prerequisites
    Import-Module -Name "ReverseDSC" -Force

    $Global:spFarmAccount = Get-Credential -Message "Credentials with Farm Admin Rights" -UserName $env:USERDOMAIN\$env:USERNAME
    Save-Credentials $Global:spFarmAccount.UserName

    $Script:spCentralAdmin = Get-SPWebApplication -IncludeCentralAdministration | Where-Object{$_.DisplayName -like '*Central Administration*'}
    $spFarm = Get-SPFarm
    $spServers = $spFarm.Servers | Where-Object{$_.Role -ne 'Invalid'}
    if($Standalone)
    {
        $i = 0;
        foreach($spServer in $spServers)
        {
            if($i -eq 0)
            {
                $spServers = @($spServer)
            }
            $i++
        }
    }
    $Script:dscConfigContent += "<# Generated with ReverseDSC " + $script:version + " #>`r`n"

    Write-Host "Scanning Operating System Version..." -BackgroundColor DarkGreen -ForegroundColor White
    Read-OperatingSystemVersion

    Write-Host "Scanning SQL Server Version..." -BackgroundColor DarkGreen -ForegroundColor White
    Read-SQLVersion

    Write-Host "Scanning Patch Levels..." -BackgroundColor DarkGreen -ForegroundColor White
    Read-SPProductVersions

    $Script:configName = "SPFarmConfig"
    if($Standalone)
    {
        $Script:configName = "SPStandalone"
    }
    if($Script:ExtractionModeValue -eq 3)
    {
        $Script:configName += "-Full"
    }
    elseif($Script:ExtractionModeValue -eq 1)
    {
        $Script:configName += "-Lite"
    }
    $Script:dscConfigContent += "Configuration $Script:configName`r`n"
    $Script:dscConfigContent += "{`r`n"
    $Script:dscConfigContent += "    <# Credentials #>`r`n"

    Write-Host "Configuring Dependencies..." -BackgroundColor DarkGreen -ForegroundColor White
    Set-Imports

    $serverNumber = 1
    $nodeLoopDone = $false;
    foreach($spServer in $spServers)
    {
        $Script:currentServerName = $spServer.Name

        <## SQL servers are returned by Get-SPServer but they have a Role of 'Invalid'. Therefore we need to ignore these. The resulting PowerShell DSC Configuration script does not take into account the configuration of the SQL server for the SharePoint Farm at this point in time. We are activaly working on giving our users an experience that is as painless as possible, and are planning on integrating the SQL DSC Configuration as part of our feature set. #>
        if($spServer.Role -ne "Invalid")
        {
            Add-ConfigurationDataEntry -Node $Script:currentServerName -Key "ServerNumber" -Value $serverNumber -Description ""

            if($serverNumber -eq 1)
            {
                $Script:dscConfigContent += "`r`n    Node `$AllNodes.Where{`$_.ServerNumber -eq '1'}.NodeName`r`n    {`r`n"
            }
            elseif(!$nodeLoopDone){
                $Script:dscConfigContent += "`r`n    Node `$AllNodes.Where{`$_.ServerNumber -ne '1'}.NodeName`r`n    {`r`n"
            }

            <# Extract the ServerRole property for SP2016 servers; #>
            $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
            $currentServer = Get-SPServer $Script:currentServerName
            if($spMajorVersion -ge 16 -and $null -eq (Get-ConfigurationDataEntry -Node $Script:currentServerName -Key "ServerRole"))
            {
                Add-ConfigurationDataEntry -Node $Script:currentServerName -Key "ServerRole" -Value $currentServer.Role -Description "MinRole for the current server;"
            }

            if($serverNumber -eq 1 -or !$nodeLoopDone)
            {
                Write-Host "["$spServer.Name"] Generating the SharePoint Prerequisites Installation..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPInstallPrereqs

                Write-Host "["$spServer.Name"] Generating the SharePoint Binary Installation..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPInstall

                Write-Host "["$spServer.Name"] Scanning the SharePoint Farm..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPFarm -ServerName $spServer.Address
            }

            if($serverNumber -eq 1)
            {
                Write-Host "["$spServer.Name"] Scanning Managed Account(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPManagedAccounts

                Write-Host "["$spServer.Name"] Scanning Web Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWebApplications

                Write-Host "["$spServer.Name"] Scanning Web Application(s) Permissions..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWebAppPermissions

                Write-Host "["$spServer.Name"] Scanning Alternate Url(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAlternateUrl

                Write-Host "["$spServer.Name"] Scanning Managed Path(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPManagedPaths

                Write-Host "["$spServer.Name"] Scanning Application Pool(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPServiceApplicationPools

                if(!$SkipSitesAndWebs)
                {
                    Write-Host "["$spServer.Name"] Scanning Content Database(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPContentDatabase

                    Write-Host "["$spServer.Name"] Scanning Quota Template(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPQuotaTemplate

                    Write-Host "["$spServer.Name"] Scanning Site Collection(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPSitesAndWebs
                }

                Write-Host "["$spServer.Name"] Scanning Diagnostic Logging Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-DiagnosticLoggingSettings

                Write-Host "["$spServer.Name"] Scanning Usage Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPUsageServiceApplication

                Write-Host "["$spServer.Name"] Scanning Web Application Policy..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWebAppPolicy

                Write-Host "["$spServer.Name"] Scanning State Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-StateServiceApplication

                Write-Host "["$spServer.Name"] Scanning User Profile Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPUserProfileServiceApplication

                Write-Host "["$spServer.Name"] Scanning Machine Translation Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPMachineTranslationServiceApp

                Write-Host "["$spServer.Name"] Cache Account(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-CacheAccounts

                Write-Host "["$spServer.Name"] Scanning Secure Store Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SecureStoreServiceApplication

                Write-Host "["$spServer.Name"] Scanning Business Connectivity Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-BCSServiceApplication

                Write-Host "["$spServer.Name"] Scanning Search Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SearchServiceApplication

                Write-Host "["$spServer.Name"] Scanning Managed Metadata Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-ManagedMetadataServiceApplication

                Write-Host "["$spServer.Name"] Scanning Access Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAccessServiceApp

                Write-Host "["$spServer.Name"] Scanning Access Services 2010 Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAccessServices2010

                Write-Host "["$spServer.Name"] Scanning Antivirus Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAntivirusSettings

                Write-Host "["$spServer.Name"] Scanning App Catalog Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAppCatalog

                Write-Host "["$spServer.Name"] Scanning Subscription Settings Service Application Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPSubscriptionSettingsServiceApp

                Write-Host "["$spServer.Name"] Scanning App Domain Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAppDomain

                Write-Host "["$spServer.Name"] Scanning App Management Service App Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAppManagementServiceApp

                Write-Host "["$spServer.Name"] Scanning App Store Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPAppStoreSettings

                Write-Host "["$spServer.Name"] Scanning Blob Cache Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPBlobCacheSettings

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Database(s) Availability Group Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPDatabaseAAG
                }

                Write-Host "["$spServer.Name"] Scanning Distributed Cache Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPDistributedCacheService

                Write-Host "["$spServer.Name"] Scanning Excel Services Application Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPExcelServiceApp

                Write-Host "["$spServer.Name"] Scanning Farm Administrator(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPFarmAdministrators

                Write-Host "["$spServer.Name"] Scanning Farm Solution(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPFarmSolution

                if($Script:ExtractionModeValue -eq 3)
                {
                    Write-Host "["$spServer.Name"] Scanning Health Rule(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPHealthAnalyzerRuleState
                }

                Write-Host "["$spServer.Name"] Scanning IRM Settings(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPIrmSettings

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Office Online Binding(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPOfficeOnlineServerBinding
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Crawl Rules(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPSearchCrawlRule
                }

                if($Script:ExtractionModeValue -eq 3)
                {
                    Write-Host "["$spServer.Name"] Scanning Search File Type(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPSearchFileType
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Search Result Source(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPSearchResultSource
                }

                Write-Host "["$spServer.Name"] Scanning Search Topology..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPSearchTopology

                Write-Host "["$spServer.Name"] Scanning Search Index Partition(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPSearchIndexPartition

                Write-Host "["$spServer.Name"] Scanning Word Automation Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWordAutomationServiceApplication

                Write-Host "["$spServer.Name"] Scanning Visio Graphics Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPVisioServiceApplication

                Write-Host "["$spServer.Name"] Scanning Work Management Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWorkManagementServiceApplication

                Write-Host "["$spServer.Name"] Scanning Performance Point Service Application..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPPerformancePointServiceApplication

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Applications Workflow Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebAppWorkflowSettings
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Applications Throttling Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebAppThrottlingSettings
                }

                if($Script:ExtractionModeValue -eq 3)
                {
                    Write-Host "["$spServer.Name"] Scanning the Timer Job States..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPTimerJobState
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Applications Usage and Deletion Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebAppSiteUseAndDeletion
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Applications Proxy Groups..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebAppProxyGroup
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Applications Extension(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebApplicationExtension
                }

                Write-Host "["$spServer.Name"] Scanning Web Applications App Domain(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWebApplicationAppDomain

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Web Application(s) General Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPWebAppGeneralSettings
                }

                Write-Host "["$spServer.Name"] Scanning Web Application(s) Blocked File Types..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPWebAppBlockedFileTypes

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning User Profile Section(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPUserProfileSection
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning User Profile Properties..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPUserProfileProperty
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning User Profile Permissions..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPUserProfileServiceAppPermissions
                }
                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning User Profile Sync Connections..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPUserProfileSyncConnection
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Trusted Identity Token Issuer(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPTrustedIdentityTokenIssuer
                }

                Write-Host "["$spServer.Name"] Scanning Farm Property Bag..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPFarmPropertyBag

                Write-Host "["$spServer.Name"] Scanning Session State Service..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPSessionStateService

                Write-Host "["$spServer.Name"] Scanning Published Service Application(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                Read-SPPublishServiceApplication

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Remote Farm Trust(s)..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPRemoteFarmTrust
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Farm Password Change Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPPasswordChangeSettings
                }

                if($Script:ExtractionModeValue -ge 2)
                {
                    Write-Host "["$spServer.Name"] Scanning Service Application(s) Security Settings..." -BackgroundColor DarkGreen -ForegroundColor White
                    Read-SPServiceAppSecurity
                }
            }

            Write-Host "["$spServer.Name"] Scanning Service Instance(s)..." -BackgroundColor DarkGreen -ForegroundColor White
            if(!$Standalone -and !$nodeLoopDone)
            {
                Read-SPServiceInstance -Servers @($spServer.Name)

                $Script:dscConfigContent += "        foreach(`$ServiceInstance in `$Node.ServiceInstances)`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $Script:dscConfigContent += "            SPServiceInstance (`$ServiceInstance.Name.Replace(`" `", `"`") + `"Instance`")`r`n"
                $Script:dscConfigContent += "            {`r`n"
                $Script:dscConfigContent += "                Name = `$ServiceInstance.Name;`r`n"
                $Script:dscConfigContent += "                Ensure = `$ServiceInstance.Ensure;`r`n"

                $Script:dscConfigContent += "                PsDscRunAsCredential = `$Creds" + ($Global:spFarmAccount.Username.Split('\'))[1].Replace("-","_").Replace(".", "_").Replace("@","").Replace(" ","") + "`r`n"

                $Script:dscConfigContent += "            }`r`n"
                $Script:dscConfigContent += "        }`r`n"
            }
            else
            {
                $servers = Get-SPServer | Where-Object{$_.Role -ne 'Invalid'}

                $serverAddresses = @()
                foreach($server in $servers)
                {
                    $serverAddresses += $server.Address
                }
                if(!$serviceLoopDone)
                {
                    Read-SPServiceInstance -Servers $serverAddresses
                    $serviceLoopDone = $true
                }
            }

            Write-Host "["$spServer.Name"] Configuring Local Configuration Manager (LCM)..." -BackgroundColor DarkGreen -ForegroundColor White
            if($serverNumber -eq 1 -or !$nodeLoopDone)
            {
                if($serverNumber -gt 1)
                {
                    $nodeLoopDone = $true
                }

                Set-LCM
                $Script:dscConfigContent += "`r`n    }`r`n"
            }

            $serverNumber++
        }
    }
    $Script:dscConfigContent += "`r`n}`r`n"
    Write-Host "Configuring Credentials..." -BackgroundColor DarkGreen -ForegroundColor White
    Set-ObtainRequiredCredentials

    if(!$Azure)
    {
        $Script:dscConfigContent += "$configName -ConfigurationData .\ConfigurationData.psd1"
    }
}

function Test-Prerequisites
{
    <# Validate the PowerShell Version #>
    if($psVersionTable.PSVersion.Major -eq 4)
    {
        Write-Host "`r`nI100"  -BackgroundColor Cyan -ForegroundColor Black -NoNewline
        Write-Host "    - PowerShell v4 detected. While this script will work just fine with v4, it is highly recommended you upgrade to PowerShell v5 to get the most out of DSC"
    }
    elseif($psVersionTable.PSVersion.Major -lt 4)
    {
        Write-Host "`r`nE100"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
        Write-Host "    - We are sorry, PowerShell v3 or lower is not supported by the ReverseDSC Engine"
        exit
    }

    <# Check to see if the SharePointDSC module is installed on the machine #>
    if(Get-Command "Get-DSCModule" -EA SilentlyContinue)
    {
        $spDSCCheck = Get-DSCResource -Module "SharePointDSC" | Where-Object{$_.Version -eq $SPDSCVersion} -ErrorAction SilentlyContinue
        <# Because the SkipPublisherCheck parameter doesn't seem to be supported on Win2012R2 / PowerShell prior to 5.1, let's set whether the parameters are specified here. #>
        if (Get-Command -Name Install-Module -ParameterName SkipPublisherCheck -ErrorAction SilentlyContinue)
        {
            $skipPublisherCheckParameter = @{SkipPublisherCheck = $true}
        }
        else
        {
            $skipPublisherCheckParameter = @{}
        }
        if($spDSCCheck.Length -eq 0)
        {
            $cmd = Get-Command Install-Module
            if($psVersionTable.PSVersion.Major -ge 5 -or $cmd)
            {
                if(!$Confirm)
                {
                    $shouldInstall = 'y'
                }
                else
                {
                    $shouldInstall = Read-Host "The SharePointDSC module could not be found on the machine. Do you wish to download and install it (y/n)?"
                }

                if($shouldInstall.ToLower() -eq "y")
                {
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                    Install-Module SharePointDSC -RequiredVersion $SPDSCVersion -Confirm:$false @skipPublisherCheckParameter
                }
                else
                {
                    Write-Host "`r`nE101"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                    Write-Host "   - We are sorry, but the script cannot continue without the SharePoint DSC module installed."
                    exit
                }
            }
            else
            {
                Write-Host "`r`nW101"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                Write-Host "   - We could not find the PackageManagement modules on the machine. Please make sure you download and install it at https://www.microsoft.com/en-us/download/details.aspx?id=51451 before executing this script"
                $Script:SPDSCPath = $moduleObject[0].Module.Path.ToLower().Replace("sharepointdsc.psd1", "").Replace("\", "/")
            }
        }
    }
    else
    {
        <# PowerShell v4 is most likely present, without the PackageManagement module. We need to manually check to see if the SharePoint
           DSC Module is present on the machine. #>
        $cmd = Get-Command Install-Module -EA SilentlyContinue
        if(!$cmd)
        {
            Write-Host "`r`nW102"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
            Write-Host "   - We could not find the PackageManagement modules on the machine. Please make sure you download and install it at https://www.microsoft.com/en-us/download/details.aspx?id=51451 before executing this script"
        }
        $moduleObject = Get-DSCResource | Where-Object{$_.Module -like "SharePointDsc"} -ErrorAction SilentlyContinue
        if(!$moduleObject)
        {
            Write-Host "`r`nE103"  -BackgroundColor Red -ForegroundColor Black -NoNewline
            Write-Host "    - Could not find the SharePointDSC Module Resource on the current server."
            exit;
        }
        $Script:SPDSCPath = $moduleObject[0].Module.Path.ToLower().Replace("sharepointdsc.psd1", "").Replace("\", "/")
    }
}

function Read-OperatingSystemVersion
{
    $servers = Get-SPServer
    $Script:dscConfigContent += "<#`r`n    Operating Systems in this Farm`r`n-------------------------------------------`r`n"
    $Script:dscConfigContent += "    Products and Language Packs`r`n"
    $Script:dscConfigContent += "-------------------------------------------`r`n"
    foreach($spServer in $servers)
    {
        $serverName = $spServer.Name
        try
        {
            $osInfo = Get-CimInstance Win32_OperatingSystem  -ComputerName $serverName -ErrorAction SilentlyContinue| Select-Object @{Label="OSName"; Expression={$_.Name.Substring($_.Name.indexof("W"),$_.Name.indexof("|")-$_.Name.indexof("W"))}} , Version ,OSArchitecture -ErrorAction SilentlyContinue
            $Script:dscConfigContent += "    [" + $serverName + "]: " + $osInfo.OSName + "(" + $osInfo.OSArchitecture + ")    ----    " + $osInfo.Version + "`r`n"
        }
        catch
        {
            $Script:ErrorLog += "[Operating System]" + $spServer.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
    $Script:dscConfigContent += "#>`r`n`r`n"
}
function Read-SQLVersion
{
    $uniqueServers = @()
    $sqlServers = Get-SPDatabase | Select-Object Server -Unique
    foreach($sqlServer in $sqlServers)
    {
        $serverName = $sqlServer.Server.Address

        if($serverName -eq $null)
        {
            $serverName = $sqlServer.Server
        }

        if(!($uniqueServers -contains $serverName))
        {
            try
            {
                $sqlVersionInfo = Invoke-SQL -Server $serverName -dbName "Master" -sqlQuery "SELECT @@VERSION AS 'SQLVersion'"
                $uniqueServers += $serverName.ToString()
                $Script:dscConfigContent += "<#`r`n    SQL Server Product Versions Installed on this Farm`r`n-------------------------------------------`r`n"
                $Script:dscConfigContent += "    Products and Language Packs`r`n"
                $Script:dscConfigContent += "-------------------------------------------`r`n"
                $Script:dscConfigContent += "    [" + $serverName.ToUpper() + "]: " + $sqlVersionInfo.SQLversion.Split("`n")[0] + "`r`n#>`r`n`r`n"
            }
            catch
            {
                $Script:ErrorLog += "[SQL Server]" + $serverName + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
    }
}


<## This function ensures all required DSC Modules are properly loaded into the current PowerShell session. #>
function Set-Imports
{
    $Script:dscConfigContent += "    Import-DscResource -ModuleName `"PSDesiredStateConfiguration`"`r`n"
    $Script:dscConfigContent += "    Import-DscResource -ModuleName `"SharePointDSC`""

    if($PSVersionTable.PSVersion.Major -eq 5)
    {
        $Script:dscConfigContent += " -ModuleVersion `"" + $SPDSCVersion + "`""
    }
    $Script:dscConfigContent += "`r`n"
}

<## This function really is optional, but helps provide valuable information about the various software components installed in the current SharePoint farm (i.e. Cummulative Updates, Language Packs, etc.). #>
function Read-SPProductVersions
{
    $Script:dscConfigContent += "<#`r`n    SharePoint Product Versions Installed on this Farm`r`n-------------------------------------------`r`n"
    $Script:dscConfigContent += "    Products and Language Packs`r`n"
    $Script:dscConfigContent += "-------------------------------------------`r`n"

    if($PSVersionTable.PSVersion -like "2.*")
    {
        $RegLoc = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $Programs = $RegLoc | where-object { $_.PsPath -like "*\Office*" } | ForEach-Object {Get-ItemProperty $_.PsPath}

        foreach($program in $Programs)
        {
            $Script:dscConfigContent += "    " +  $program.DisplayName + " -- " + $program.DisplayVersion + "`r`n"
        }
    }
    else
    {
        $regLoc = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $programs = $regLoc | where-object { $_.PsPath -like "*\Office*" } | ForEach-Object {Get-ItemProperty $_.PsPath}
        $components = $regLoc | where-object { $_.PsPath -like "*1000-0000000FF1CE}" } | ForEach-Object {Get-ItemProperty $_.PsPath}

        foreach($program in $programs)
        {
            $productCodes = $_.ProductCodes
            $component = @() + ($components |     where-object { $_.PSChildName -in $productCodes } | ForEach-Object {Get-ItemProperty $_.PsPath})
            foreach($component in $components)
            {
                $Script:dscConfigContent += "    " + $component.DisplayName + " -- " + $component.DisplayVersion + "`r`n"
            }
        }
    }
    $Script:dscConfigContent += "#>`r`n"
}
function Read-SPInstall
{
    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "FullInstallation" -Value "`$True" -Description "Specifies whether or not the DSC configuration script will install the SharePoint Prerequisites and Binaries;"
    $Script:dscConfigContent += "        if(`$ConfigurationData.NonNodeData.FullInstallation)`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $Script:dscConfigContent += "            SPInstall BinaryInstallation" + "`r`n            {`r`n"
    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SPInstallationBinaryPath" -Value "\\<location>" -Description "Location of the SharePoint Binaries (local path or network share);"
    $Script:dscConfigContent += "                BinaryDir = `$ConfigurationData.NonNodeData.SPInstallationBinaryPath;`r`n"
    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SPProductKey" -Value "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" -Description "SharePoint Product Key"
    $Script:dscConfigContent += "                ProductKey = `$ConfigurationData.NonNodeData.SPProductKey;`r`n"
    $Script:dscConfigContent += "                Ensure = `"Present`";`r`n"

    $Script:dscConfigContent += "                PSDscRunAsCredential = `$Creds" + ($Global:spFarmAccount.Username.Split('\'))[1].Replace("-","_").Replace(".", "_") + ";`r`n"
    $Script:dscConfigContent += "            }`r`n"
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPInstallPrereqs
{
    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "FullInstallation" -Value "`$True" -Description "Specifies whether or not the DSC configuration script will install the SharePoint Prerequisites and Binaries;"
    $Script:dscConfigContent += "        if(`$ConfigurationData.NonNodeData.FullInstallation)`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $Script:dscConfigContent += "            SPInstallPrereqs PrerequisitesInstallation" + "`r`n            {`r`n"
    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SPPrereqsInstallerPath" -Value "\\<location>" -Description "Location of the SharePoint Prerequisites Installer .exe (Local path or Network Share);"
    $Script:dscConfigContent += "                InstallerPath = `$ConfigurationData.NonNodeData.SPPrereqsInstallerPath;`r`n"
    $Script:dscConfigContent += "                OnlineMode = `$True;`r`n"
    $Script:dscConfigContent += "                Ensure = `"Present`";`r`n"

    $Script:dscConfigContent += "                PSDscRunAsCredential = `$Creds" + ($Global:spFarmAccount.Username.Split('\'))[1].Replace("-","_").Replace(".", "_") + ";`r`n"

    $Script:dscConfigContent += "            }`r`n"
    $Script:dscConfigContent += "        }`r`n"
}

<## This function declares the SPFarm object required to create the config and admin database for the resulting SharePoint Farm. #>
function Read-SPFarm (){
    param(
        [string]$ServerName,
        [bool]$RunCentralAdmin
    )
    $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarm\MSFT_SPFarm.psm1")
    Import-Module $module

    $Script:dscConfigContent += "        SPFarm " + [System.Guid]::NewGuid().ToString() + "`r`n        {`r`n"
    $params = Get-DSCFakeParameters -ModulePath $module
    <# If not SP2016, remove the server role param. #>
    if ($spMajorVersion -ne 16)
    {
        $params.Remove("ServerRole")
    }

    <# Can't have both the InstallAccount and PsDscRunAsCredential variables present. Remove InstallAccount if both are there. #>
    if($params.Contains("InstallAccount"))
    {
        $params.Remove("InstallAccount")
    }

    <# WA - Bug in 1.6.0.0 Get-TargetResource does not return the current Authentication Method; #>
    $caAuthMethod = "NTLM"
    if(!$Script:spCentralAdmin.IisSettings[0].DisableKerberos)
    {
        $caAuthMethod = "Kerberos"
    }
    $params.CentralAdministrationAuth = $caAuthMethod
    $params.CentralAdministrationPort = $Script:spCentralAdmin.IisSettings[0].ServerBindings.Port
    $params.FarmAccount = $Global:spFarmAccount
    $params.Passphrase = $Global:spFarmAccount
    $results = Get-TargetResource @params

    <# Remove the default generated PassPhrase and ensure the resulting Configuration Script will prompt user for it. #>
    $results.Remove("Passphrase");

    <# WA - Bug in 1.6.0.0 Get-TargetResource not returning name if aliases are used #>
    $configDB = Get-SPDatabase | Where-Object{$_.GetType().Name -eq "SPConfigurationDatabase"}
    $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

    if($null -eq (Get-ConfigurationDataEntry -Key "DatabaseServer"))
    {
        Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $configDB.NormalizedDataSource -Description "Name of the Database Server associated with the destination SharePoint Farm;"
    }

    if($null -eq (Get-ConfigurationDataEntry -Key "PassPhrase"))
    {
        Add-ConfigurationDataEntry -Node "NonNodeData" -Key "PassPhrase" -Value "pass@word1" -Description "SharePoint Farm's PassPhrase;"
    }

    $Script:dscConfigContent += "            Passphrase = New-Object System.Management.Automation.PSCredential ('Passphrase', (ConvertTo-SecureString -String `$ConfigurationData.NonNodeData.PassPhrase -AsPlainText -Force));`r`n"

    if(!$results.ContainsKey("RunCentralAdmin"))
    {
        $results.Add("RunCentralAdmin", $RunCentralAdmin)
    }

    if($spMajorVersion -ge 16)
    {
        if(!$results.Contains("ServerRole"))
        {
            $results.Add("ServerRole", "`$Node.ServerRole")
        }
        else
        {
            $results["ServerRole"] = "`$Node.ServerRole"
        }
    }
    else
    {
        $results.Remove("ServerRole")
    }
    $results = Repair-Credentials -results $results
    $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
    if($spMajorVersion -ge 16)
    {
        $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "ServerRole"
    }
    $Script:dscConfigContent += $currentBlock
    $Script:dscConfigContent += "        }`r`n"

    <# SPFarm Feature Section #>
    if($Script:ExtractionModeValue -eq 3)
    {
        $versionFilter = $spMajorVersion.ToString() + "*"
        $farmFeatures = Get-SPFeature | Where-Object{$_.Scope -eq "Farm" -and $_.Version -like $versionFilter}
        $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
        Import-Module $moduleFeature
        $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

        $featuresAlreadyAdded = @()
        $i = 1
        $total = $farmFeatures.Length
        foreach($farmFeature in $farmFeatures)
        {
            try
            {
                $featureName = $farmFeature.DisplayName
                Write-Host "Scanning Farm Feature [$i/$total] {$featureName}"
                if(!$featuresAlreadyAdded.Contains($farmFeature.DisplayName))
                {
                    $featuresAlreadyAdded += $featureName
                    $paramsFeature.Name = $featureName
                    $paramsFeature.FeatureScope = "Farm"
                    $resultsFeature = Get-TargetResource @paramsFeature

                    if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
                    {
                        $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                        $Script:dscConfigContent += "        {`r`n"

                        <# Manually add the InstallAccount param due to a bug in 1.6.0.0 that returns a param named InstalAccount (typo) instead.
                        https://github.com/PowerShell/SharePointDsc/issues/481 #>
                        if($resultsFeature.ContainsKey("InstalAccount"))
                        {
                            $resultsFeature.Remove("InstalAccount")
                        }

                        $resultsFeature = Repair-Credentials -results $resultsFeature
                        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsFeature -ModulePath $moduleFeature
                        $Script:dscConfigContent += "        }`r`n"
                    }
                }
                $i++
            }
            catch
            {
                $Script:ErrorLog += "[Farm Features]" + $farmFeature.DisplayName + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
    }
}

<## This function obtains a reference to every Web Application in the farm and declares their properties (i.e. Port, Associated IIS Application Pool, etc.). #>
function Read-SPWebApplications (){
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebApplication\MSFT_SPWebApplication.psm1")
    Import-Module $module
    $spWebApplications = Get-SPWebApplication | Sort-Object -Property Name
    $params = Get-DSCFakeParameters -ModulePath $module

    $i = 1;
    $total = $spWebApplications.Length
    foreach($spWebApp in $spWebApplications)
    {
        try
        {
            $webAppName = $params.Name = $spWebApp.Name
            Write-Host "Scanning SPWebApplication [$i/$total] {$webAppName}"
            Import-Module $module
            $Script:dscConfigContent += "        SPWebApplication " + $spWebApp.Name.Replace(" ", "") + "`r`n        {`r`n"

            $params.Name = $webAppName
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results

            $appPoolAccount = Get-Credentials $results.ApplicationPoolAccount
            if($null -eq $appPoolAccount)
            {
                Save-Credentials -UserName $results.ApplicationPoolAccount
            }
            $results.ApplicationPoolAccount = (Resolve-Credentials -UserName $results.ApplicationPoolAccount) + ".UserName"

            if($null -eq $results.Get_Item("AllowAnonymous"))
            {
                $results.Remove("AllowAnonymous")
            }

            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
            $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

            $currentDSCBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentDSCBlock = Convert-DSCStringParamToVariable -DSCBlock $currentDSCBlock -ParameterName "ApplicationPoolAccount"
            $currentDSCBlock = Convert-DSCStringParamToVariable -DSCBlock $currentDSCBlock -ParameterName "DatabaseServer"
            $Script:dscConfigContent += $currentDSCBlock
            $Script:dscConfigContent += "        }`r`n"

            if($Script:ExtractionModeValue -ge 2)
            {
                Write-Host "    -> Scanning SharePoint Designer Settings"
                Read-SPDesignerSettings($spWebApplications.Url.ToString(), "WebApplication", $spWebApp.Name.Replace(" ", ""))
            }

            <# SPWebApplication Feature Section #>
            if($Script:ExtractionModeValue -eq 3)
            {
                $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
                $versionFilter = $spMajorVersion.ToString() + "*"
                $webAppFeatures = Get-SPFeature | Where-Object{$_.Scope -eq "WebApplication" -and $_.Version -like $versionFilter}
                $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
                Import-Module $moduleFeature
                $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

                $j = 1
                $totalFeat = $webAppFeatures.Length
                foreach($webAppFeature in $webAppFeatures)
                {
                    try
                    {
                        $displayName = $webAppFeature.DisplayName
                        Write-Host "    -> Scanning Feature [$j/$totalFeat] {$displayName}"
                        $paramsFeature.Name = $displayName
                        $paramsFeature.FeatureScope = "WebApplication"
                        $paramsFeature.Url = $spWebApp.Url
                        $resultsFeature = Get-TargetResource @paramsFeature

                        if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
                        {
                            $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                            $Script:dscConfigContent += "        {`r`n"

                            <# Manually add the InstallAccount param due to a bug in 1.6.0.0 that returns a param named InstalAccount (typo) instead.
                            https://github.com/PowerShell/SharePointDsc/issues/481 #>
                            if($resultsFeature.ContainsKey("InstalAccount"))
                            {
                                $resultsFeature.Remove("InstalAccount")
                            }
                            $resultsFeature = Repair-Credentials -results $resultsFeature
                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsFeature -ModulePath $moduleFeature
                            $Script:dscConfigContent += "            DependsOn = `"[SPWebApplication]" + $spWebApp.Name.Replace(" ", "") + "`";`r`n"
                            $Script:dscConfigContent += "        }`r`n"
                        }
                        $j++
                    }
                    catch
                    {
                        $Script:ErrorLog += "[Web Application Feature]" + $webAppFeature.DisplayName + "`r`n"
                        $Script:ErrorLog += "$_`r`n`r`n"
                    }
                }
            }

            <# Outgoing Email Setting Region #>
            $moduleEmail = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPOutgoingEmailSettings\MSFT_SPOutgoingEmailSettings.psm1")
            Import-Module $moduleEmail
            $paramsEmail = Get-DSCFakeParameters -ModulePath $moduleEmail

            $paramsEmail.WebAppUrl = $spWebApp.Url
            $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
            if($spMajorVersion.ToString() -eq "15" -and $paramsEmail.Contains("UseTLS"))
            {
                $paramsEmail.Remove("UseTLS")
            }
            if($spMajorVersion.ToString() -eq "15" -and $paramsEmail.Contains("SMTPPort"))
            {
                $paramsEmail.Remove("SMTPPort")
            }

            $resultsEmail = Get-TargetResource @paramsEmail
            if($null -eq $resultsEmail["SMTPPort"])
            {
                $resultsEmail.Remove("SMTPPort")
            }
            if($null -eq $resultsEmail["UseTLS"])
            {
                $resultsEmail.Remove("UseTLS")
            }
            if($null -eq $resultsEmail["ReplyToAddress"])
            {
                $resultsEmail["ReplyToAddress"] = "*"
            }
            if($null -ne $resultsEmail["SMTPServer"] -and "" -ne $resultsEmail["SMTPServer"])
            {
                Write-Host "    -> Scanning Outgoing Email Settings"
                $Script:dscConfigContent += "        SPOutgoingEmailSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $resultsEmail = Repair-Credentials -results $resultsEmail
                if($null -eq $resultsEmail.ReplyToAddress -or $resultsEmail.ReplyToAddress -eq "")
                {
                    $resultsEmail.ReplyToAddress = "*"
                }
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsEmail -ModulePath $moduleEmail
                $Script:dscConfigContent += "            DependsOn = `"[SPWebApplication]" + $spWebApp.Name.Replace(" ", "") + "`";`r`n"
                $Script:dscConfigContent += "        }`r`n"
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Web Application]" + $spWebApp.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Repair-Credentials($results)
{
    if($null -ne $results)
    {
        <## Cleanup the InstallAccount param first (even if we may be adding it back) #>
        if($null -ne $results.ContainsKey("InstallAccount"))
        {
            $results.Remove("InstallAccount")
        }

        if($null -ne $results.ContainsKey("PsDscRunAsCredential"))
        {
            $results.Remove("PsDscRunAsCredential")
        }

        $results.Add("PsDscRunAsCredential", "`$Creds" + ($Global:spFarmAccount.Username.Split('\'))[1].Replace("-","_").Replace(".", "_").Replace("@","").Replace(" ",""))

        return $results
    }
    return $null
}

<## This function loops through every IIS Application Pool that are associated with the various existing Service Applications in the SharePoint farm. ##>
function Read-SPServiceApplicationPools
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPServiceAppPool\MSFT_SPServiceAppPool.psm1")
    Import-Module $module
    $spServiceAppPools = Get-SPServiceApplicationPool | Sort-Object -Property Name

    $params = Get-DSCFakeParameters -ModulePath $module

    $i = 1
    $total = $spServiceAppPools.Length
    foreach($spServiceAppPool in $spServiceAppPools)
    {
        try
        {
            $appPoolName = $spServiceAppPool.Name
            Write-Host "Scanning SPServiceApplicationPool [$i/$total] {$appPoolName}"
            $Script:dscConfigContent += "        SPServiceAppPool " + $spServiceAppPool.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $appPoolName
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results

            $serviceAccount = Get-Credentials $results.ServiceAccount
            if($null -eq $serviceAccount)
            {
                Save-Credentials -UserName $results.ServiceAccount
            }
            $results.ServiceAccount = (Resolve-Credentials -UserName $results.ServiceAccount) + ".UserName"

            if($null -eq $results.Get_Item("AllowAnonymous"))
            {
                $results.Remove("AllowAnonymous")
            }
            $currentDSCBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentDSCBlock = Convert-DSCStringParamToVariable -DSCBlock $currentDSCBlock -ParameterName "ServiceAccount"
            $Script:dscConfigContent += $currentDSCBlock

            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Service Application Pool]" + $spServiceAppPool.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPQuotaTemplate()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPQuotaTemplate\MSFT_SPQuotaTemplate.psm1")
    Import-Module $module
    $contentService = Get-SPDSCContentService

    $params = Get-DSCFakeParameters -ModulePath $module

    $quotaGUID = ""
    $i = 1
    $total = $contentservice.QuotaTemplates.Count
    foreach($quota in $contentservice.QuotaTemplates)
    {
        try
        {
            $quotaName = $quota.Name
            Write-Host "Scanning Quota Template [$i/$total] {$quotaName}"
            $quotaGUID = [System.Guid]::NewGuid().ToString()
            $Script:DH_SPQUOTATEMPLATE.Add($quotaName, $quotaGUID)

            $Script:dscConfigContent += "        SPQuotaTemplate " + $quotaGUID + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $quota.Name
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Quota Template]" + $quota.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves a list of all site collections, no matter what Web Application they belong to. The Url attribute helps the xSharePoint DSC Resource determine what Web Application they belong to. #>
function Read-SPSitesAndWebs ()
{
    $spSites = Get-SPSite -Limit All
    $siteGuid = $null
    $siteTitle = $null
    $dependsOnItems = @()
    $sc = Get-SPDSCContentService

    $i = 1
    $total = $spSites.Length
    foreach($spSite in $spSites)
    {
        try
        {
            if(!$spSite.IsSiteMaster)
            {
                $siteTitle = $spSite.RootWeb.Title
                $siteUrl = $spSite.Url
                Write-Host "Scanning SPSite [$i/$total] {$siteUrl}"

                $dependsOnItems = @("[SPWebApplication]" + $spSite.WebApplication.Name.Replace(" ", ""))
                $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSite\MSFT_SPSite.psm1")
                Import-Module $module
                $params = Get-DSCFakeParameters -ModulePath $module
                $siteGuid = [System.Guid]::NewGuid().toString()
                $siteTitle = $spSite.RootWeb.Title
                if($siteTitle -eq $null)
                {
                    $siteTitle = "SiteCollection"
                }
                $Script:dscConfigContent += "        SPSite " + $siteGuid + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.Url = $spSite.Url
                $results = Get-TargetResource @params

                <# WA - Somehow the WebTemplateID returned for App Catalog is 18, but the template is APPCATALOG#0 #>
                if($results.Template -eq "APPCATALOG#18")
                {
                    $results.Template = "APPCATALOG#0"
                }
                <# If the current Quota ID is 0, it means no quota templates were used. Remove param in that case. #>
                if($spSite.Quota.QuotaID -eq 0)
                {
                    $results.Remove("QuotaTemplate")
                }
                else
                {
                    $quotaTemplateName = $sc.QuotaTemplates | Where-Object{$_.QuotaId -eq $spsite.Quota.QuotaID}
                    if($null -ne $quotaTemplateName -and $null -ne $quotaTemplateName.Name)
                    {
                        if($Script:DH_SPQUOTATEMPLATE.ContainsKey($quotaTemplateName.Name))
                        {
                            $dependsOnItems += "[SPQuotaTemplate]" + $Script:DH_SPQUOTATEMPLATE.Item($quotaTemplateName.Name)
                        }
                    }
                    else
                    {
                        $results.Remove("QuotaTemplate")
                    }
                }
                if(!$results.Get_Item("SecondaryOwnerAlias"))
                {
                    $results.Remove("SecondaryOwnerAlias")
                }
                if(!$results.Get_Item("SecondaryEmail"))
                {
                    $results.Remove("SecondaryEmail")
                }
                if(!$results.Get_Item("OwnerEmail"))
                {
                    $results.Remove("OwnerEmail")
                }
                if(!$results.Get_Item("HostHeaderWebApplication"))
                {
                    $results.Remove("HostHeaderWebApplication")
                }
                if(!$results.Get_Item("Name"))
                {
                    $results.Remove("Name")
                }
                if(!$results.Get_Item("Description"))
                {
                    $results.Remove("Description")
                }
                else
                {
                    $results.Description = $results.Description.Replace("`"", "'").Replace("`r`n", ' `
                    ')
                }
                $dependsOnClause = Get-DSCDependsOnBlock($dependsOnItems)
                $results = Repair-Credentials -results $results

                $ownerAlias = Get-Credentials -UserName $results.OwnerAlias
                $plainTextUser = $false;
                if(!$ownerAlias)
                {
                    $plainTextUser = $true
                    $ownerAlias = $results.OwnerAlias
                }
                $currentBlock = ""
                if($null -ne $ownerAlias -and !$plainTextUser)
                {
                    $results.OwnerAlias = (Resolve-Credentials -UserName $results.OwnerAlias) + ".UserName"
                }

                if($results.ContainsKey("SecondaryOwnerAlias"))
                {
                    $secondaryOwner = Get-Credentials -UserName $results.SecondaryOwnerAlias
                    if($null -ne $secondaryOwner)
                    {
                        $results.SecondaryOwnerAlias = (Resolve-Credentials -UserName $results.SecondaryOwnerAlias) + ".UserName"
                    }
                    else
                    {
                        $secondaryOwner = $results.SecondaryOwnerAlias
                    }
                }
                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                if($null -ne $results.SecondaryOwnerAlias -and $results.SecondaryOwnerAlias.StartsWith("`$"))
                {
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "SecondaryOwnerAlias"
                }
                if($null -ne $results.OwnerAlias -and $results.OwnerAlias.StartsWith("`$"))
                {
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "OwnerAlias"
                }
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "            DependsOn =  " + $dependsOnClause + "`r`n"
                $Script:dscConfigContent += "        }`r`n"

                <# Nik20170112 - There are restrictions preventing this setting from being applied if the PsDscRunAsCredential parameter is not used.
                            Since this is only available in WMF 5, we check to see if the node farm we are extracting the configuration from is
                            running at least PowerShell v5 before reading the Site Collection level SPDesigner settings. #>
                if($PSVersionTable.PSVersion.Major -ge 5 -and $Script:ExtractionModeValue -ge 2)
                {
                    Read-SPDesignerSettings($spSite.Url, "SiteCollection")
                }

                Read-SPSiteUrl($spSite.Url)
                if($Script:ExtractionModeValue -eq 3)
                {
                    $webs = Get-SPWeb -Limit All -Site $spsite
                    $j = 1
                    $totalWebs = $webs.Length
                    foreach($spweb in $webs)
                    {
                        try
                        {
                            $webUrl = $spweb.Url
                            Write-Host "    -> Scanning Web [$j/$totalWebs] {$webUrl}"
                            $moduleWeb = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWeb\MSFT_SPWeb.psm1")
                            Import-Module $moduleWeb
                            $paramsWeb = Get-DSCFakeParameters -ModulePath $moduleWeb
                            $paramsWeb.Url = $webUrl
                            $resultsWeb = Get-TargetResource @paramsWeb
                            $resultsWeb.Description = $resultsWeb.Description.Replace("`"", "'").Replace("`r`n", ' `
                            ')
                            $Script:dscConfigContent += "        SPWeb " + [System.Guid]::NewGuid().toString() + "`r`n"
                            $Script:dscConfigContent += "        {`r`n"
                            $resultsWeb = Repair-Credentials -results $resultsWeb
                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsWeb -ModulePath $moduleWeb
                            $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
                            $Script:dscConfigContent += "        }`r`n"

                            <# SPWeb Feature Section #>
                            if($Script:ExtractionModeValue -eq 3)
                            {
                                $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
                                $versionFilter = $spMajorVersion.ToString() + "*"
                                $webFeatures = Get-SPFeature | Where-Object{$_.Scope -eq "Web" -and $_.Version -like $versionFilter}
                                $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
                                Import-Module $moduleFeature
                                $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

                                $k = 0
                                $totalWebFeat = $webFeatures.Length
                                foreach($webFeature in $webFeatures)
                                {
                                    try
                                    {
                                        $webFeatureName = $webFeature.DisplayName
                                        Write-Host "        ---> Scanning Web Feature [$k/$totalWebFeat] {$webFeatureName}"

                                        $paramsFeature.Name = $webFeatureName
                                        $paramsFeature.FeatureScope = "Web"
                                        $paramsFeature.Url = $spWeb.Url
                                        $resultsFeature = Get-TargetResource @paramsFeature

                                        if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
                                        {
                                            $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                                            $Script:dscConfigContent += "        {`r`n"

                                            $resultsFeature = Repair-Credentials -results $resultsFeature
                                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsFeature -ModulePath $moduleFeature
                                            $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
                                            $Script:dscConfigContent += "        }`r`n"
                                        }
                                        $k++
                                    }
                                    catch
                                    {
                                        $Script:ErrorLog += "[Web Feature]" + $webFeature.DisplayName + "`r`n"
                                        $Script:ErrorLog += "$_`r`n`r`n"
                                    }
                                }
                            }
                            $j++
                        }
                        catch
                        {
                            $Script:ErrorLog += "[Web]" + $spweb.Url + "`r`n"
                            $Script:ErrorLog += "$_`r`n`r`n"
                        }
                    }
                }
                <# SPSite Feature Section #>
                if($Script:ExtractionModeValue -eq 3)
                {
                    $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
                    $versionFilter = $spMajorVersion.ToString() + "*"
                    $siteFeatures = Get-SPFeature | Where-Object{$_.Scope -eq "Site" -and $_.Version -like $versionFilter}
                    $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
                    Import-Module $moduleFeature
                    $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

                    $k = 0
                    $totalSiteFeat = $siteFeatures.Length
                    foreach($siteFeature in $siteFeatures)
                    {
                        try
                        {
                            $siteFeatName = $siteFeature.DisplayName
                            Write-Host "    -> Scanning SPSite feature [$k/$totalSiteFeat] {$siteFeatName}"

                            $paramsFeature.Name = $siteFeatName
                            $paramsFeature.FeatureScope = "Site"
                            $paramsFeature.Url = $spSite.Url
                            $resultsFeature = Get-TargetResource @paramsFeature

                            if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
                            {
                                $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                                $Script:dscConfigContent += "        {`r`n"

                                <# Manually add the InstallAccount param due to a bug in 1.6.0.0 that returns a param named InstalAccount (typo) instead.
                                https://github.com/PowerShell/SharePointDsc/issues/481 #>
                                if($resultsFeature.ContainsKey("InstalAccount"))
                                {
                                    $resultsFeature.Remove("InstalAccount")
                                }
                                $resultsFeature = Repair-Credentials -results $resultsFeature
                                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsFeature -ModulePath $moduleFeature
                                $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
                                $Script:dscConfigContent += "        }`r`n"
                            }
                            $k++
                        }
                        catch
                        {
                            $Script:ErrorLog += "[Site Collection Feature]" + $siteFeature.DisplayName + "`r`n"
                            $Script:ErrorLog += "$_`r`n`r`n"
                        }
                    }
                }
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Site Collection]" + $spSite.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function generates a list of all Managed Paths, no matter what their associated Web Application is. The xSharePoint DSC Resource uses the WebAppUrl attribute to identify what Web Applicaton they belong to. #>
function Read-SPManagedPaths
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPManagedPath\MSFT_SPManagedPath.psm1")
    Import-Module $module

    $spWebApps = Get-SPWebApplication
    $params = Get-DSCFakeParameters -ModulePath $module

    foreach($spWebApp in $spWebApps)
    {
        $spManagedPaths = Get-SPManagedPath -WebApplication $spWebApp.Url | Sort-Object -Property Name

        $i = 1
        $total = $spManagedPaths.Length
        foreach($spManagedPath in $spManagedPaths)
        {
            try
            {
                Write-Host "Scanning Web Application Managed Path [$i/$total] {"$spManagedPath.Name"}"
                if($spManagedPath.Name.Length -gt 0 -and $spManagedPath.Name -ne "sites")
                {
                    $Script:dscConfigContent += "        SPManagedPath " + [System.Guid]::NewGuid().toString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    if($spManagedPath.Name -ne $null)
                    {
                        $params.RelativeUrl = $spManagedPath.Name
                    }
                    $params.WebAppUrl = $spWebApp.Url
                    $params.HostHeader = $false;
                    if($params.Contains("InstallAccount"))
                    {
                        $params.Remove("InstallAccount")
                    }
                    $results = Get-TargetResource @params

                    $results = Repair-Credentials -results $results

                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
                $i++
            }
            catch
            {
                $Script:ErrorLog += "[Web Application Managed Path]" + $spManagedPath.Name + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
    }
    $spManagedPaths = Get-SPManagedPath -HostHeader | Sort-Object -Property Name
    $i = 0
    $total = $spManagedPaths.Length
    foreach($spManagedPath in $spManagedPaths)
    {
        try
        {
            Write-Host "Scanning Host Header Managed Path [$i/$total] {"$spManagedPath.Name"}"
            if($spManagedPath.Name.Length -gt 0 -and $spManagedPath.Name -ne "sites")
            {
                $Script:dscConfigContent += "        SPManagedPath " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"

                if($spManagedPath.Name -ne $null)
                {
                    $params.RelativeUrl = $spManagedPath.Name
                }
                if($params.ContainsKey("Explicit"))
                {
                    $params.Explicit = ($spManagedPath.Type -eq "ExplicitInclusion")
                }
                else
                {
                    $params.Add("Explicit", ($spManagedPath.Type -eq "ExplicitInclusion"))
                }
                $params.WebAppUrl = "*"
                $params.HostHeader = $true;
                $results = Get-TargetResource @params
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Host Header Managed Path]" + $spManagedPath.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves all Managed Accounts in the SharePoint Farm. The Account attribute sets the associated credential variable (each managed account is declared as a variable and the user is prompted to Manually enter the credentials when first executing the script. See function "Set-ObtainRequiredCredentials" for more details on how these variales are set. #>
function Read-SPManagedAccounts()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPManagedAccount\MSFT_SPManagedAccount.psm1")
    Import-Module $module
    $managedAccounts = Get-SPManagedAccount

    $i = 1
    $total = $managedAccounts.Length
    foreach($managedAccount in $managedAccounts)
    {
        try
        {
            $mAccountName = $managedAccount.UserName
            Write-Host "Scanning SPManagedAccount [$i/$total] {$mAccountName}"

            $Script:dscConfigContent += "        SPManagedAccount " + [System.Guid]::NewGuid().toString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            <# WA - 1.6.0.0 has a bug where the Get-TargetResource returns an array of all ManagedAccount (see Issue #533) #>
            $schedule = $null
            if($null -ne $managedAccount.ChangeSchedule)
            {
                $schedule = $managedAccount.ChangeSchedule.ToString()
            }
            $results = @{AccountName = $managedAccount.UserName; EmailNotification = $managedAccount.DaysBeforeChangeToEmail; PreExpireDays = $managedAccount.DaysBeforeExpiryToChange;Schedule = $schedule;Ensure="Present";}
            $results["Account"] = "`$Creds" + ($managedAccount.UserName.Split('\'))[1]
            $results = Repair-Credentials -results $results

            $accountName = Get-Credentials -UserName $managedAccount.UserName
            if($null -eq $accountName)
            {
                Save-Credentials -UserName $managedAccount.UserName
            }
            $results.AccountName = (Resolve-Credentials -UserName $managedAccount.UserName) + ".UserName"

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "AccountName"
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Managed Account]" + $managedAccount.UserName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves all Services in the SharePoint farm. It does not care if the service is enabled or not. It lists them all, and simply sets the "Ensure" attribute of those that are disabled to "Absent". #>
function Read-SPServiceInstance($Servers)
{
    $servicesMasterList = @()
    foreach($Server in $Servers)
    {
        Write-Host "Scanning SPServiceInstance on {$Server}"
        $serviceInstancesOnCurrentServer = Get-SPServiceInstance | Where-Object{$_.Server.Name -eq $Server} | Sort-Object -Property TypeName
        $serviceStatuses = @()
        $ensureValue = "Present"

        $i = 1
        $total = $serviceInstancesOnCurrentServer.Length
        foreach($serviceInstance in $serviceInstancesOnCurrentServer)
        {
            try
            {
                $serviceTypeName = $serviceInstance.TypeName
                Write-Host "    -> Scanning instance [$i/$total] {$serviceTypeName}"

                if($serviceInstance.Status -eq "Online")
                {
                    $ensureValue = "Present"
                }
                else
                {
                    $ensureValue = "Absent"
                }

                $currentService = @{Name = $serviceInstance.TypeName; Ensure = $ensureValue}

                if($serviceInstance.TypeName -ne "Distributed Cache" -and $serviceInstance.TypeName -ne "User Profile Synchronization Service")
                {
                    $serviceStatuses += $currentService
                }
                if($ensureValue -eq "Present" -and !$servicesMasterList.Contains($serviceInstance.TypeName))
                {
                    $servicesMasterList += $serviceTypeName
                    Write-Verbose $serviceTypeName
                    if($serviceTypeName -eq "Distributed Cache")
                    {
                        # Do Nothing - This is handled by its own call later on.
                    }
                    elseif($serviceTypeName -eq "User Profile Synchronization Service")
                    {
                        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileSyncService\MSFT_SPUserProfileSyncService.psm1")
                        Import-Module $module
                        $params = Get-DSCFakeParameters -ModulePath $module
                        $params.Ensure = $ensureValue
                        $params.FarmAccount = $Global:spFarmAccount
                        if($null -eq $params.InstallAccount)
                        {
                            $params.Remove("InstallAccount")
                        }
                        $results = Get-TargetResource @params
                        if($ensureValue -eq "Present")
                        {
                            $Script:dscConfigContent += "        SPUserProfileSyncService " + $serviceTypeName.Replace(" ", "") + "Instance`r`n"
                            $Script:dscConfigContent += "        {`r`n"

                            if($results.Contains("InstallAccount"))
                            {
                                $results.Remove("InstallAccount")
                            }
                            if(!$results.Contains("FarmAccount"))
                            {
                                $results.Add("FarmAccount", $Global:spFarmAccount)
                            }
                            $results = Repair-Credentials -results $results
                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                            $Script:dscConfigContent += "        }`r`n"
                        }
                    }
                }
                $i++
            }
            catch
            {
                $Script:ErrorLog += "[Service Instance]" + $serviceInstance.TypeName + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
        Add-ConfigurationDataEntry -Node $Server -Key "ServiceInstances" -Value $serviceStatuses
    }
}

<## This function retrieves all settings related to Diagnostic Logging (ULS logs) on the SharePoint farm. #>
function Read-DiagnosticLoggingSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDiagnosticLoggingSettings\MSFT_SPDiagnosticLoggingSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $Script:dscConfigContent += "        SPDiagnosticLoggingSettings ApplyDiagnosticLogSettings`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $results = Repair-Credentials -results $results

    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "LogPath" -Value $results.LogPath -Description "Path where the SharePoint ULS logs will be stored;"
    $results.LogPath = "`$ConfigurationData.NonNodeData.LogPath"

    $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "LogPath"
    $Script:dscConfigContent += $currentBlock
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPSiteURL($siteUrl)
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSiteURL\MSFT_SPSiteURL.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $results = Get-TargetResource @params

    if($results.Intranet -ne $null -or $results.Internet -ne $null -or $results.Custom -ne $null -or $results.Extranet -ne $null)
    {
        $blockGUID = New-Guid
        $Script:dscConfigContent += "        SPSiteUrl " + $blockGUID.ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Url = $siteUrl
        $results = Repair-Credentials -results $results

        $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPMachineTranslationServiceApp()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPMachineTranslationServiceApp\MSFT_SPMachineTranslationServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $machineTranslationServiceApps = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Machine Translation Service"}
    $i = 1
    $total = $machineTranslationServiceApps.Length
    foreach($machineTranslation in $machineTranslationServiceApps)
    {
        try
        {
            $serviceName = $machineTranslation.Name
            Write-Host "Scanning Machine Translation Service [$i/$total] {$serviceName}"

            $Script:dscConfigContent += "        SPMachineTranslationServiceApp " + [System.Guid]::NewGuid().toString()  + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $serviceName
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results

            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
            $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Machine Translation Service Application]" + $machineTranslation.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppPolicy()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppPolicy\MSFT_SPWebAppPolicy.psm1")
    Import-Module $module

    $webApps = Get-SPWebApplication

    $i = 1
    $total = $webApps.Length
    foreach($webApp in $webApps)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
        $webAppUrl = $webApp.Url
        Write-Host "Scanning Web App Policies [$i/$total] {$webAppUrl}"

        $params.WebAppUrl = $webAppUrl
        $Script:dscConfigContent += "        SPWebAppPolicy " + [System.Guid]::NewGuid().toString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $fake = New-CimInstance -ClassName Win32_Process -Property @{Handle=0} -Key Handle -ClientOnly

        if(!$params.Contains("Members"))
        {
            $params.Add("Members", $fake);
        }
        $results = Get-TargetResource @params

        if($null -ne $results.Members)
        {
            $newMembers = @()
            foreach($member in $results.Members)
            {
                if($member.UserName.Contains("\"))
                {
                    $resultPermission = Get-SPWebPolicyPermissions -params $member
                    $newMembers += $resultPermission
                }
            }
            $params.Members = $newMembers
        }

        if($null -eq $results.MembersToExclude)
        {
            $results.Remove("MembersToExclude")
        }

        if($null -eq $results.MembersToInclude)
        {
            $results.Remove("MembersToInclude")
        }

        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
        $i++
    }
}

<## This function retrieves all settings related to the SharePoint Usage Service Application, assuming it exists. #>
function Read-SPUsageServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUsageApplication\MSFT_SPUsageApplication.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $usageApplication = Get-SPUsageApplication
    if($usageApplication.Length -gt 0)
    {
        $Script:dscConfigContent += "        SPUsageApplication " + $usageApplication.TypeName.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $usageApplication.Name
        $params.Ensure = "Present"
        $results = Get-TargetResource @params
        $results.DatabaseCredentials = $Global:spFarmAccount
        $failOverFound = $false

        $results = Repair-Credentials -results $results
        if($null -eq $results.FailOverDatabaseServer)
        {
            $results.Remove("FailOverDatabaseServer")
        }
        else
        {
            $failOverFound = $true
            Add-ConfigurationDataEntry -Node $env:COMPUTERNAME -Key "UsageAppFailOverDatabaseServer" -Value $results.FailOverDatabaseServer -Description "Name of the Usage Service Application Failover Database;"
            $results.FailOverDatabaseServer = "`$ConfigurationData.NonNodeData.UsageAppFailOverDatabaseServer"
        }
        $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"
        Add-ConfigurationDataEntry -Node "NonNodeData" -Key "UsageLogLocation" -Value $results.UsageLogLocation -Description "Path where the Usage Logs will be stored;"
        $results.UsageLogLocation = "`$ConfigurationData.NonNodeData.UsageLogLocation"

        $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "UsageLogLocation"
        $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"

        if($failOverFound)
        {
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "FailOverDatabaseServer"
        }

        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
}

<## This function retrieves settings associated with the State Service Application, assuming it exists. #>
function Read-StateServiceApplication ($modulePath, $params)
{
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPStateServiceApp\MSFT_SPStateServiceApp.psm1")
        Import-Module $module
    }

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $stateApplications = Get-SPStateServiceApplication

    $i = 1
    $total = $stateApplications.Length
    foreach($stateApp in $stateApplications)
    {
        try
        {
            if($stateApp -ne $null)
            {
                $serviceName = $stateApp.DisplayName
                Write-Host "Scanning State Service Application [$i/$total] {$serviceName}"

                $params.Name = $serviceName
                $Script:dscConfigContent += "        SPStateServiceApp " + $serviceName.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params
                $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

                $results = Repair-Credentials -results $results
                $currentBlock += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[State service Application]" + $stateApp.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves information about all the "Super" accounts (Super Reader & Super User) used for caching. #>
function Read-CacheAccounts ($modulePath, $params)
{
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPCacheAccounts\MSFT_SPCacheAccounts.psm1")
        Import-Module $module
    }

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $webApps = Get-SPWebApplication

    $i = 1
    $total = $webApps.Length
    foreach($webApp in $webApps)
    {
        $webAppUrl = $webApp.Url
        Write-Host "Scanning Cache Account [$i/$total] {$webAppUrl}"

        $params.WebAppUrl = $webAppUrl
        $results = Get-TargetResource @params

        if($results.SuperReaderAlias -ne "" -and $results.SuperUserAlias -ne "")
        {
            $Script:dscConfigContent += "        SPCacheAccounts " + $webApp.DisplayName.Replace(" ", "") + "CacheAccounts`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
        $i++
    }
}

<## This function retrieves settings related to the User Profile Service Application. #>
function Read-SPUserProfileServiceApplication ($modulePath, $params)
{
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileServiceApp\MSFT_SPUserProfileServiceApp.psm1")
        Import-Module $module
    }

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $ups = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "User Profile Service Application"}

    $sites = Get-SPSite -Limit All
    if($sites.Length -gt 0)
    {
        $context = Get-SPServiceContext $sites[0]
        try
        {
            $catch = new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($context)
            $catch = $null
        }
        catch
        {
            if($null -ne $ups)
            {
                Write-Host "`r`nW103"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                Write-Host "   - Farm Account does not have Full Control on the User Profile Service Application."
            }
        }

        if($ups -ne $null)
        {
            $i = 1
            $total = $ups.Length
            foreach($upsInstance in $ups)
            {
                try
                {
                    $serviceName = $upsInstance.DisplayName
                    Write-Host "Scanning User Profile Service Application [$i/$total] {$serviceName}"

                    $params.Name = $serviceName
                    $currentBlock = "        SPUserProfileServiceApp " + [System.Guid]::NewGuid().toString() + "`r`n"
                    $currentBlock += "        {`r`n"

                    if($null -eq $params.InstallAccount)
                    {
                        $params.Remove("InstallAccount")
                    }

                    $results = Get-TargetResource @params
                    if($results.Contains("MySiteHostLocation") -and $results.Get_Item("MySiteHostLocation") -eq "*")
                    {
                        $results.Remove("MySiteHostLocation")
                    }

                    if($results.Contains("InstallAccount"))
                    {
                        $results.Remove("InstallAccount")
                    }
                    $results = Repair-Credentials -results $results

                    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SyncDBServer" -Value $results.SyncDBServer -Description "Name of the User Profile Service Sync Database Server;"
                    $results.SyncDBServer = "`$ConfigurationData.NonNodeData.SyncDBServer"

                    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "ProfileDBServer" -Value $results.ProfileDBServer -Description "Name of the User Profile Service Profile Database Server;"
                    $results.ProfileDBServer = "`$ConfigurationData.NonNodeData.ProfileDBServer"

                    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SocialDBServer" -Value $results.SocialDBServer -Description "Name of the User Profile Social Database Server;"
                    $results.SocialDBServer = "`$ConfigurationData.NonNodeData.SocialDBServer"

                    $currentBlock += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "SyncDBServer"
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "ProfileDBServer"
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "SocialDBServer"
                    $Script:dscConfigContent += $currentBlock
                    $Script:dscConfigContent += "        }`r`n"
                    $i++
                }
                catch
                {
                    $Script:ErrorLog += "[User Profile Service Application]" + $upsInstance.DisplayName + "`r`n"
                    $Script:ErrorLog += "$_`r`n`r`n"
                }
            }
        }
    }
}

<## This function retrieves all settings related to the Secure Store Service Application. Currently this function makes a direct call to the Secure Store database on the farm's SQL server to retrieve information about the logging details. There are currently no publicly available hooks in the SharePoint/Office Server Object Model that allow us to do it. This forces the user executing this reverse DSC script to have to install the SQL Server Client components on the server on which they execute the script, which is not a "best practice". #>
<# TODO: Change the logic to extract information about the logging from being a direct SQL call to something that uses the Object Model. #>
function Read-SecureStoreServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSecureStoreServiceApp\MSFT_SPSecureStoreServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Secure Store Service Application"}

    $i = 1
    $total = $ssas
    foreach($ssa in $ssas)
    {
        try
        {
            $serviceName = $ssa.DisplayName
            Write-Host "Scanning Secure Store Service Application [$i/$total] {$serviceName}"

            $params.Name = $serviceName
            $Script:dscConfigContent += "        SPSecureStoreServiceApp " + $ssa.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params

            <# WA - Issue with 1.6.0.0 where DB Aliases not returned in Get-TargetResource #>
            $secStoreDBs = Get-SPDatabase | Where-Object{$_.Type -eq "Microsoft.Office.SecureStoreService.Server.SecureStoreServiceDatabase"}
            $results.DatabaseName = $secStoreDBs.DisplayName
            $results.DatabaseServer = $secStoreDBs.NormalizedDataSource

            <# WA - Can't dynamically retrieve value from the Secure Store at the moment #>
            if(!$results.Contains("AuditingEnabled"))
            {
                $results.Add("AuditingEnabled", $true)
            }

            if($results.Contains("InstallAccount"))
            {
                $results.Remove("InstallAccount")
            }

            $results = Repair-Credentials -results $results

            $foundFailOver = $false
            if($null -eq $results.FailOverDatabaseServer)
            {
                $results.Remove("FailOverDatabaseServer")
            }
            else
            {
                $foundFailOver = $true
                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SecureStoreFailOverDatabaseServer" -Value $results.FailOverDatabaseServer -Description "Name of the SQL Server that hosts the FailOver database for your SharePoint Farm's Secure Store Service Application;"
                $results.FailOverDatabaseServer = "`$ConfigurationData.NonNodeData.SecureStoreFailOverDatabaseServer"
            }

            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
            $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
            if($foundFailOver)
            {
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "FailOverDatabaseServer"
            }
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Secure Store Service Application]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves settings related to the Managed Metadata Service Application. #>
function Read-ManagedMetadataServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPManagedMetadataServiceApp\MSFT_SPManagedMetadataServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $mms = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Managed Metadata Service"}
    if (Get-Command "Get-SPMetadataServiceApplication" -errorAction SilentlyContinue)
    {
        $i = 1
        $total = $mms.Length
        foreach($mmsInstance in $mms)
        {
            try
            {
                if($mmsInstance -ne $null)
                {
                    $serviceName = $mmsInstance.Name
                    Write-Host "Scanning Managed Metadata Service [$i/$total] {$serviceName}"

                    $params.Name = $serviceName
                    $Script:dscConfigContent += "        SPManagedMetaDataServiceApp " + $serviceName.Replace(" ", "") + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $results = Get-TargetResource @params

                    <# WA - Issue with 1.6.0.0 where DB Aliases not returned in Get-TargetResource #>
                    $results["DatabaseServer"] = CheckDBForAliases -DatabaseName $results["DatabaseName"]
                    $results = Repair-Credentials -results $results

                    if(!$results.Languages)
                    {
                        $results.Remove("Languages")
                    }

                    $results.TermStoreAdministrators = Set-TermStoreAdministratorsBlock $results.TermStoreAdministrators

                    Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
                    $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

                    $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $currentBlock = Set-TermStoreAdministrators $currentBlock
                    $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
                    $Script:dscConfigContent += $currentBlock
                    $Script:dscConfigContent += "        }`r`n"
                }
                $i++
            }
            catch
            {
                $Script:ErrorLog += "[Managed Metadata Service Application]" + $mmsInstance.DisplayName + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
    }
}

function Set-TermStoreAdministrators($DSCBlock)
{
    $newLine = "TermStoreAdministrators = @("

    $startPosition = $DSCBlock.IndexOf("TermStoreAdministrators = @")
    if($startPosition -ge 0)
    {
        $endPosition = $DSCBlock.IndexOf("`r`n", $startPosition)
        if($endPosition -ge 0)
        {
            $DSCLine = $DSCBlock.Substring($startPosition, $endPosition - $startPosition)
            $originalLine = $DSCLine
            $DSCLine = $DSCLine.Replace("TermStoreAdministrators = @(","").Replace(");","").Replace(" ","")
            $members = $DSCLine.Split(',')

            $i = 1
            $total = $members.Length
            foreach($member in $members)
            {
                Write-Host "    -> Scanning TermStore Admins [$i/$total]"
                if($member.StartsWith("`"`$"))
                {
                    $newLine += $member.Replace("`"","") + ", "
                }
                else
                {
                    $newLine += $member + ", "
                }
                $i++
            }
            if($newLine.EndsWith(", "))
            {
                $newLine = $newLine.Remove($newLine.Length - 2, 2)
            }
            $newLine += ");"
            $DSCBlock = $DSCBlock.Replace($originalLine, $newLine)
        }
    }

    return $DSCBlock
}

function Set-TermStoreAdministratorsBlock($TermStoreAdminsLine)
{
    $newArray = @()
    foreach($admin in $TermStoreAdminsLine)
    {
        if(!($admin -like "BUILTIN*"))
        {
            $account = Get-Credentials -UserName $admin
            if($null -eq $account)
            {
                Save-Credentials -UserName $admin
            }
            $newArray += (Resolve-Credentials -UserName $admin) + ".UserName"
        }
        else
        {
            $newArray += $admin
        }
    }
    return $newArray
}

function Read-SPWordAutomationServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWordAutomationServiceApp\MSFT_SPWordAutomationServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $was = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Word Automation Services"}

    $i = 1
    $total = $was.Length
    foreach($wa in $was)
    {
        try
        {
            if($wa -ne $null)
            {
                $serviceName = $wa.Name
                Write-Host "Scanning Word Automation Service Application [$i/$total] {$serviceName}"

                $params.Name = $serviceName
                $Script:dscConfigContent += "        SPWordAutomationServiceApp " + $serviceName.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
                $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Managed Metadata Service Application]" + $wa.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPVisioServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPVisioServiceApp\MSFT_SPVisioServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $was = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Visio Graphics Service Application"}

    $i = 1
    $total = $was.Length
    foreach($wa in $was)
    {
        try
        {
            if($wa -ne $null)
            {
                $serviceName = $wa.Name
                Write-Host "Scanning Visio Service Application [$i/$total] {$serviceName}"

                $params.Name = $serviceName
                $Script:dscConfigContent += "        SPVisioServiceApp " + $wa.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Visio Graphics Service Application]" + $wa.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPTrustedIdentityTokenIssuer()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPTrustedIdentityTokenIssuer\MSFT_SPTrustedIdentityTokenIssuer.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $tips = Get-SPTrustedIdentityTokenIssuer

    $i = 1
    $total = $tips.Length
    foreach($tip in $tips)
    {
        try
        {
            $tokenName = $tip.Name
            Write-Host "Scanning Trusted Identity Token Issuer [$i/$total] {$tokenName}"

            $params.Name = $tokenName
            $params.Description = $tip.Description
            $Script:dscConfigContent += "        SPTrustedIdentityTokenIssuer " + [System.Guid]::NewGuid().toString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $fake = New-CimInstance -ClassName Win32_Process -Property @{Handle=0} -Key Handle -ClientOnly

            if(!$params.Contains("ClaimsMappings"))
            {
                $params.Add("ClaimsMappings", $fake)
            }
            $results = Get-TargetResource @params

            foreach($ctm in $results.ClaimsMappings)
            {
                $ctmResult = Get-SPClaimTypeMapping -params $ctm
                $results.ClaimsMappings = $ctmResult
            }
            if($null -ne $results.Get_Item("SigningCertificateThumbprint") -and $results.Contains("SigningCertificateFilePath"))
            {
                $results.Remove("SigningCertificateFilePath")
            }

            if($results.Contains("InstallAccount"))
            {
                $results.Remove("InstallAccount")
            }
            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Trusted Identity Token Issuer]" + $tip.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWorkManagementServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWorkManagementServiceApp\MSFT_SPWorkManagementServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $was = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Work Management Service Application"}
    foreach($wa in $was)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.Name = $wa.Name
                $Script:dscConfigContent += "        SPWorkManagementServiceApp " + $wa.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Work Management Service Application]" + $wa.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPTimerJobState
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPTimerJobState\MSFT_SPTimerJobState.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $spTimers = Get-SPTimerJob
    $totalTimers = $spTimers.Length
    $i = 0;
    foreach($timer in $spTimers)
    {
        try
        {
            $i++
            Write-Host "Scanning Timer Job {"$timer.Name"}[$i/$totalTimers]..."
            if($timer -ne $null -and $timer.TypeName -ne "Microsoft.SharePoint.Administration.Health.SPHealthAnalyzerJobDefinition")
            {
                $params.TypeName = $timer.TypeName
                if($null -ne $timer.WebApplication)
                {
                    $params.WebAppUrl = $timer.WebApplication.Url;
                }
                else
                {
                    $params.WebAppUrl = "N/A";
                }

                <# TODO: Remove comment tags when version 2.0.0.0 of SharePointDSC gets released;#>
                $Script:dscConfigContent += "<#`r`n"
                $Script:dscConfigContent += "        SPTimerJobState " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
                $Script:dscConfigContent += "#>`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Timer Job]" + $timer.TypeName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPPerformancePointServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPPerformancePointServiceApp\MSFT_SPPerformancePointServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $was = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "PerformancePoint Service Application"}
    foreach($wa in $was)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.Name = $wa.Name
                $Script:dscConfigContent += "        SPPerformancePointServiceApp " + $wa.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
                $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"

                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[PerformancePoint Service Application]" + $wa.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppWorkflowSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppWorkflowSettings\MSFT_SPWebAppWorkflowSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.Url = $wa.Url
                $Script:dscConfigContent += "        SPWebAppWorkflowSettings " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Workflow Settings]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppThrottlingSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppThrottlingSettings\MSFT_SPWebAppThrottlingSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.Url = $wa.Url
                $Script:dscConfigContent += "        SPWebAppThrottlingSettings " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results.HappyHour = Get-SPWebAppHappyHour -params $results.HappyHour
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Throttling Settings]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppSiteUseAndDeletion()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppSiteUseAndDeletion\MSFT_SPWebAppSiteUseAndDeletion.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.Url = $wa.Url
                $Script:dscConfigContent += "        SPWebAppSiteUseAndDeletion " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Site Use and Deletion]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebApplicationExtension()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebApplicationExtension\MSFT_SPWebApplicationExtension.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $zones = @("Default","Intranet","Internet","Extranet","Custom")
    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.WebAppUrl = $wa.Url

                for($i = 0; $i -lt $zones.Length; $i++)
                {
                    if($null -ne $wa.IisSettings[$zones[$i]])
                    {
                        $params.Zone = $zones[$i]
                        $Script:dscConfigContent += "        SPWebApplicationExtension " + [System.Guid]::NewGuid().toString() + "`r`n"
                        $Script:dscConfigContent += "        {`r`n"
                        $results = Get-TargetResource @params

                        if($results.Contains("InstallAccount"))
                        {
                            $results.Remove("InstallAccount")
                        }
                        if("" -eq $results.HostHeader)
                        {
                            $results.Remove("HostHeader")
                        }
                        if($null -eq $results.AuthenticationProvider)
                        {
                            $results.Remove("AuthenticationProvider")
                        }
                        $results = Repair-Credentials -results $results
                        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                        $Script:dscConfigContent += "        }`r`n"
                    }
                }
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Extensions]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppPermissions()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppPermissions\MSFT_SPWebAppPermissions.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try {
            if($wa -ne $null)
            {
                $params.WebAppUrl = $wa.Url
                $params.Remove("ListPermissions")
                $params.Remove("SitePermissions")
                $params.Remove("PersonalPermissions")
                $Script:dscConfigContent += "        SPWebAppPermissions " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }

                <# Fix an issue with SP DSC (forward) 1.6.0.0 #>
                if($results.WebAppUrl -eq "url")
                {
                    $results.WebAppUrl = $wa.Url
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Permissions]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebAppProxyGroup()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppProxyGroup\MSFT_SPWebAppProxyGroup.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($wa in $webApps)
    {
        try
        {
            if($wa -ne $null)
            {
                $params.WebAppUrl = $wa.Url
                $params.ServiceAppProxyGroup = $wa.ServiceApplicationProxyGroup.FriendlyName
                $Script:dscConfigContent += "        SPWebAppProxyGroup " + [System.Guid]::NewGuid().toString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Web Application Proxy Group]" + $wa.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<## This function retrieves settings related to the Business Connectivity Service Application. #>
function Read-BCSServiceApplication ($modulePath, $params)
{
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPBCSServiceApp\MSFT_SPBCSServiceApp.psm1")
        Import-Module $module
    }

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $bcsa = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Business Data Connectivity Service Application"}

    foreach($bcsaInstance in $bcsa)
    {
        try
        {
            if($bcsaInstance -ne $null)
            {
                $Script:dscConfigContent += "        SPBCSServiceApp " + $bcsaInstance.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.Name = $bcsaInstance.DisplayName
                $results = Get-TargetResource @params

                <# WA - Issue with 1.6.0.0 where DB Aliases not returned in Get-TargetResource #>
                $results["DatabaseServer"] = CheckDBForAliases -DatabaseName $results["DatabaseName"]

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }
                $results = Repair-Credentials -results $results

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
                $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"
                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Business Connectivity Service Application]" + $bcsaInstance.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function CheckDBForAliases()
{
    param(
      [string]$DatabaseName
    )

    $dbServer = Get-SPDatabase | Where-Object{$_.Name -eq $DatabaseName}
    return $dbServer.NormalizedDataSource
}

<## This function retrieves settings related to the Search Service Application. #>
function Read-SearchServiceApplication()
{
    $searchSA = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Search Service Application"}

    $i = 1
    $total = $searchSA.Length
    foreach($searchSAInstance in $searchSA)
    {
        try
        {
            if($searchSAInstance -ne $null)
            {
                $serviceName = $searchSAInstance.Name
                Write-Host "Scanning Search Service Application [$i/$total] {$serviceName}"
                $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchServiceApp\MSFT_SPSearchServiceApp.psm1")
                Import-Module $module
                $params = Get-DSCFakeParameters -ModulePath $module

                $Script:dscConfigContent += "        SPSearchServiceApp " + $searchSAInstance.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.Name = $serviceName
                $params.ApplicationPool = $searchSAInstance.ApplicationPool.Name
                $results = Get-TargetResource @params
                if($results.Get_Item("CloudIndex") -eq $false)
                {
                    $results.Remove("CloudIndex")
                }

                if($results.Contains("InstallAccount"))
                {
                    $results.Remove("InstallAccount")
                }

                if($null -eq $results.SearchCenterUrl)
                {
                    $results.Remove("SearchCenterUrl")
                }

                <# Nik20170111 - Fix a bug in 1.5.0.0 where DatabaseName and DatabaseServer is not properly returned #>
                $results["DatabaseName"] = $searchSAInstance.SearchAdminDatabase.Name
                $results["DatabaseServer"] = $searchSAInstance.SearchAdminDatabase.Server.Name
                $results = Repair-Credentials -results $results

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
                $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"

                #region Search Content Sources
                $moduleContentSource = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchContentSource\MSFT_SPSearchContentSource.psm1")
                Import-Module $moduleContentSource
                $paramsContentSource = Get-DSCFakeParameters -ModulePath $moduleContentSource
                $contentSources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $searchSAInstance.Name

                $j = 1
                $totalCS = $contentSources.Length
                foreach($contentSource in $contentSources)
                {
                    $csName = $contentSource.Name
                    Write-Host "    -> Scanning Content Source [$j/$totalCS] {$csName}"

                    $sscsGuid = [System.Guid]::NewGuid().toString()

                    $paramsContentSource.Name = $csName
                    $paramsContentSource.ServiceAppName  = $serviceName

                    if(!$source.Type -eq "CustomRepository")
                    {
                        $Script:dscConfigContent += "        SPSearchContentSource " + $contentSource.Name.Replace(" ", "") + $sscsGuid + "`r`n"
                        $Script:dscConfigContent += "        {`r`n"

                        $resultsContentSource = Get-TargetResource @paramsContentSource

                        $searchScheduleModulePath = Resolve-Path ($Script:SPDSCPath + "\Modules\SharePointDsc.Search\SPSearchContentSource.Schedules.psm1")
                        Import-Module -Name $searchScheduleModulePath
                        # TODO: Figure out way to properly pass CimInstance objects and then add the schedules back;
                        $incremental = Get-SPDSCSearchCrawlSchedule -Schedule $contentSource.IncrementalCrawlSchedule
                        $full = Get-SPDSCSearchCrawlSchedule -Schedule $contentSource.FullCrawlSchedule

                        $resultsContentSource.IncrementalSchedule = Get-SPCrawlSchedule $incremental
                        $resultsContentSource.FullSchedule = Get-SPCrawlSchedule $full

                        $resultsContentsource = Repair-Credentials -results $resultsContentSource

                        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $resultsContentSource -ModulePath $moduleContentSource
                        $Script:dscConfigContent += "        }`r`n"
                    }
                    $j++
                }
                #endregion
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Search Service Application]" + $searchSAInstance.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Get-SPCrawlSchedule($params)
{
    $currentSchedule = "MSFT_SPSearchCrawlSchedule{`r`n"
    foreach($key in $params.Keys)
    {
        try
        {
            $currentSchedule += "                " + $key + " = `"" + $params[$key] + "`"`r`n"
        }
        catch
        {
            $Script:ErrorLog += "[MSFT_SPSearchCrawlSchedule]" + $key + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
    $currentSchedule += "            }"
    return $currentSchedule
}

function Get-SPServiceAppSecurityMembers($member)
{
    try
    {
        $catch = [System.Guid]::Parse($member.UserName)
        $isUserGuid = $true
    }
    catch
    {
        $isUserGuid = $false
    }

    if($member.AccessLevel -ne $null -and !($member.AccessLevel -match "^[\d\.]+$") -and (!$isUserGuid) -and $member.AccessLevel -ne "")
    {
        $userName = Get-Credentials -UserName $member.UserName
        if($null -eq $userName)
        {
            Save-Credentials -UserName $member.UserName
        }
        return "MSFT_SPServiceAppSecurityEntry {`
            Username    = " + (Resolve-Credentials -UserName $member.UserName) + ".UserName;`
            AccessLevel = `"" + $member.AccessLevel + "`";`
        }"
    }
    return $null
}

function Get-SPWebPolicyPermissions($params)
{
    $permission = "MSFT_SPWebPolicyPermissions{`r`n"
    foreach($key in $params.Keys)
    {
        try
        {
            $isCredentials = $false
            if($key.ToLower() -eq "username")
            {
                if(!($params[$key].ToUpper() -like "NT AUTHORITY*"))
                {
                    $memberUserName = Get-Credentials -UserName $params[$key]
                    if($null -eq $memberUserName)
                    {
                        Save-Credentials -UserName $params[$key]
                    }
                    $isCredentials = $true
                }
            }

            if(($params[$key].ToString().ToLower() -eq "false" -or $params[$key].ToString().ToLower() -eq "true") -and !$isCredentials)
            {
                $permission += "                " + $key + " = `$" + $params[$key] + "`r`n"
            }
            elseif(!$isCredentials)
            {
                $permission += "                " + $key + " = `"" + $params[$key] + "`"`r`n"
            }
            else
            {
                $permission += "                " + $key + " =  " + (Resolve-Credentials -UserName $params[$key]) + ".UserName`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[MSFT_SPWebPolicyPermissions]" + $key + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
    $permission += "            }"
    return $permission
}

function Get-SPClaimTypeMapping($params)
{
    $ctm = "MSFT_SPClaimTypeMapping{`r`n"
    foreach($key in $params.Keys)
    {
        try
        {
            if($params[$key].ToString().ToLower() -eq "false" -or $params[$key].ToString().ToLower() -eq "true")
            {
                $ctm += "                " + $key + " = `$" + $params[$key] + "`r`n"
            }
            else
            {
                $ctm += "                " + $key + " = `"" + $params[$key] + "`"`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[MSFT_SPClaimTypeMapping]" + $key + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
    $ctm += "            }"
    return $ctm
}

function Get-SPWebAppHappyHour($params)
{
    $happyHour = "MSFT_SPWebApplicationHappyHour{`r`n"
    foreach($key in $params.Keys)
    {
        try
        {
            $happyHour += "                " + $key + " = `"" + $params[$key] + "`"`r`n"
        }
        catch
        {
            $Script:ErrorLog += "[MSFT_SPWebApplicationHappyHour]" + $key + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
    $happyHour += "            }"
    return $happyHour
}

function Read-SPContentDatabase()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPContentDatabase\MSFT_SPContentDatabase.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $spContentDBs = Get-SPContentDatabase

    $i = 1
    $total = $spContentDBs.Length
    foreach($spContentDB in $spContentDBs)
    {
        try
        {
            $dbName = $spContentDB.Name
            Write-Host "Scanning Content Database [$i/$total] {$dbName}"
            $Script:dscConfigContent += "        SPContentDatabase " + $spContentDB.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $dbName
            $params.WebAppUrl = $spContentDB.WebApplication.Url
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results

            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
            $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Content Database]" + $spContentDB.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPAccessServiceApp()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAccessServiceApp\MSFT_SPAccessServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication
    $serviceApps = $serviceApps | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Access.Services.MossHost.AccessServicesWebServiceApplication"}

    $i = 1
    $total = $serviceApps.Length
    foreach($spAccessService in $serviceApps)
    {
        try
        {
            $serviceName = $spAccessService.Name
            Write-Host "Scanning Access Service Application [$i/$total] {$serviceName}"

            $params.Name = $serviceName
            $params.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"
            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results

            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
            $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"
            $Script:dscConfigContent += "        SPAccessServiceApp " + $serviceName.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Access Service Application]" + $spAccessService.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPAccessServices2010()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAccessServices2010\MSFT_SPAccessServices2010.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication
    $serviceApps = $serviceApps | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Access.Server.MossHost.AccessServerWebServiceApplication"}

    $i = 1
    $total = $serviceApps.Length
    foreach($spAccessService in $serviceApps)
    {
        try
        {
            $serviceName = $spAccessService.Name
            Write-Host "Scanning Access 2010 Service Application [$i/$total] {$serviceName}"

            $params.Name = $serviceName
            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results

            $Script:dscConfigContent += "        SPAccessServices2010 " + $spAccessService.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += $currentBlock
            $Script:dscConfigContent += "        }`r`n"
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Access 2010 Service Application]" + $spAccessService.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPAppCatalog()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppCatalog\MSFT_SPAppCatalog.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication

    foreach($webApp in $webApps)
    {
        try
        {
            $feature = $webApp.Features.Item([Guid]::Parse("f8bea737-255e-4758-ab82-e34bb46f5828"))
            if($null -ne $feature)
            {
                $appCatalogSiteId = $feature.Properties["__AppCatSiteId"].Value
                $appCatalogSite = $webApp.Sites | Where-Object{$_.ID -eq $appCatalogSiteId}

                if($null -ne $appCatalogSite)
                {
                    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppCatalog\MSFT_SPAppCatalog.psm1")
                    Import-Module $module
                    $params = Get-DSCFakeParameters -ModulePath $module

                    $catUrl = $appCatalogSite.Url
                    Write-Host "Scanning App Catalog {$catUrl}"
                    $Script:dscConfigContent += "        SPAppCatalog " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $params.SiteUrl = $catUrl
                    $results = Get-TargetResource @params
                    $results = Repair-Credentials -results $results
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
            }
        }
        catch
        {
            $Script:ErrorLog += "[App Catalog]" + $webApp.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPAppDomain()
{
    $serviceApp = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "App Management Service Application"}
    $appDomain =  Get-SPAppDomain
    if($serviceApp.Length -ge 1 -and $appDomain.Length -ge 1)
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppDomain\MSFT_SPAppDomain.psm1")
        Import-Module $module
        $params = Get-DSCFakeParameters -ModulePath $module
        $Script:dscConfigContent += "        SPAppDomain " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params
        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPSearchFileType()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchFileType\MSFT_SPSearchFileType.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}
    $i = 1
    $total = $ssas.Length

    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $serviceName = $ssa.DisplayName
                Write-Host "Scanning Search File Type for Search Application [$i/$total] {$serviceName}"
                $fileFormats = Get-SPEnterpriseSearchFileFormat -SearchApplication $ssa

                $j = 1
                $totalFT = $fileFormats.Length
                foreach($fileFormat in $fileFormats)
                {
                    $fileType = $fileFormat.Identity
                    Write-Host "    -> Scanning File Type [$j/$totalFT] {$fileType}"

                    $Script:dscConfigContent += "        SPSearchFileType " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $params.ServiceAppName = $serviceName
                    $params.FileType = $fileType

                    $results = Get-TargetResource @params

                    $results = Repair-Credentials -results $results

                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"

                    $j++
                }
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Search File Type]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSearchIndexPartition()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchIndexPartition\MSFT_SPSearchIndexPartition.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}

    $i = 1
    $total = $ssas.Length
    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $serviceName = $ssa.DisplayName
                Write-Host "Scanning Index Partitions for Search Service Application [$i/$total] {$serviceName}"

                $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $ssa
                $currentTopology = $ssa.ActiveTopology
                $indexComponents = Get-SPEnterpriseSearchComponent -SearchTopology $currentTopology | `
                    Where-Object -FilterScript {$_.GetType().Name -eq "IndexComponent"}

                [System.Collections.ArrayList]$indexesAlreadyScanned = @()
                $j = 1
                $totalIndex = $indexComponents.Length
                foreach($indexComponent in $indexComponents)
                {
                    try
                    {
                        if(!$indexesAlreadyScanned.Contains($indexComponent.IndexPartitionOrdinal))
                        {
                            $icServerName = $indexComponent.ServerName
                            Write-Host "    -> Index Component [$j/$totalIndex] {$icServerName}"

                            $indexesAlreadyScanned += $indexComponent.IndexPartitionOrdinal
                            $Script:dscConfigContent += "        SPSearchIndexPartition " + [System.Guid]::NewGuid().ToString() + "`r`n"
                            $Script:dscConfigContent += "        {`r`n"
                            $params.ServiceAppName = $serviceName
                            $params.Index = $indexComponent.IndexPartitionOrdinal
                            $params.Servers = $indexComponent.ServerName
                            $params.RootDirectory = $indexComponent.RootDirectory
                            $results = Get-TargetResource @params

                            Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchIndexPartitionServers" -Value $results.Servers -Description "List of Servers that will host the Search Index Partitions;"
                            $results.Servers = "`$ConfigurationData.NonNodeData.SearchIndexPartitionServers"

                            $results = Repair-Credentials -results $results

                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                            $Script:dscConfigContent += "        }`r`n"
                            $j++
                        }
                    }
                    catch
                    {
                        $Script:ErrorLog += "[Index Component]" + $indexComponent.ServerName + "`r`n"
                        $Script:ErrorLog += "$_`r`n`r`n"
                    }
                }
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Search Index Partition]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSearchTopology()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchTopology\MSFT_SPSearchTopology.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}

    $i = 1
    $total = $ssas.Length
    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $serviceName = $ssa.DisplayName
                Write-Host "Scanning Search Topology for Service Application [$i/$total] {$serviceName}"
                $Script:dscConfigContent += "        SPSearchTopology " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.ServiceAppName = $serviceName
                $results = Get-TargetResource @params

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchContentProcessingServers" -Value $results.ContentProcessing -Description "List of servers that will act as Search Content Processors;"
                $results.ContentProcessing = "`$ConfigurationData.NonNodeData.SearchContentProcessingServers"

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchAnalyticsProcessingServers" -Value $results.AnalyticsProcessing -Description "List of servers that will act as Search Analytics Processors;"
                $results.AnalyticsProcessing = "`$ConfigurationData.NonNodeData.SearchAnalyticsProcessingServers"

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchIndexPartitionServers" -Value $results.IndexPartition -Description "List of servers that will host the Search Index Partitions;"
                $results.IndexPartition = "`$ConfigurationData.NonNodeData.SearchIndexPartitionServers"

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchCrawlerServers" -Value $results.Crawler -Description "List of servers that will act as Search Crawlers;"
                $results.Crawler = "`$ConfigurationData.NonNodeData.SearchCrawlerServers"

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "SearchAdminServers" -Value $results.Admin -Description "List of servers that will host the Search Admin Components;"
                $results.Admin = "`$ConfigurationData.NonNodeData.SearchAdminServers"

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "QueryProcessingServers" -Value $results.QueryProcessing -Description "List of servers that will host the Search Query Components;"
                $results.QueryProcessing = "`$ConfigurationData.NonNodeData.QueryProcessingServers"

                if($results.FirstPartitionDirectory.Length -gt 1)
                {
                    $results.FirstPartitionDirectory = $results.FirstPartitionDirectory[0]
                }

                $results = Repair-Credentials -results $results

                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
                $i++
            }
        }
        catch
        {
            $Script:ErrorLog += "[Search Topology]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSearchResultSource()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchResultSource\MSFT_SPSearchResultSource.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}

    $i = 1
    $total = $ssas.Length
    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $serviceName = $ssa.DisplayName
                Write-Host "Scanning Results Sources for Search Service Application [$i/$total] {$serviceName}"
                $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $ssa
                $searchSiteUrl = $ssa.SearchCenterUrl -replace "/pages"
                $searchSite = Get-SPWeb -Identity $searchSiteUrl -ErrorAction SilentlyContinue

                if(!$null -eq $searchSite)
                {
                    $adminNamespace = "Microsoft.Office.Server.Search.Administration"
                    $objectLevel = [Microsoft.Office.Server.Search.Administration.SearchObjectLevel]
                    $searchOwner = New-Object -TypeName "$adminNamespace.SearchObjectOwner" `
                                            -ArgumentList @(
                                                $objectLevel::Ssa,
                                                $searchSite
                                            )
                    $resultSources = Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $searchOwner
                    $j = 1
                    $totalRS = $resultSources.Length
                    foreach($resultSource in $resultSources)
                    {
                        <# Filter out the hidden Local SharePoint Graph provider since it is not supported by SharePointDSC. #>
                        if($resultSource.Name -ne "Local SharePoint Graph")
                        {
                            $rsName = $resultSource.Name
                            Write-Host "    -> Scanning Results Source [$j/$totalRS] {$rsName}"
                            $Script:dscConfigContent += "        SPSearchResultSource " + [System.Guid]::NewGuid().ToString() + "`r`n"
                            $Script:dscConfigContent += "        {`r`n"
                            $params.SearchServiceAppName = $serviceName
                            $params.Name = $rsName
                            $results = Get-TargetResource @params
                            if($null -eq $results.Get_Item("ConnectionUrl"))
                            {
                                $results.Remove("ConnectionUrl")
                            }
                            $results = Repair-Credentials -results $results
                            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                            $Script:dscConfigContent += "        }`r`n"
                        }
                        $j++
                    }
                }
            }
        }
        catch
        {
            $Script:ErrorLog += "[Search Result Source]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSearchCrawlRule()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchCrawlRule\MSFT_SPSearchCrawlRule.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}
    $i = 1
    $total = $ssas.Length
    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $serviceName = $ssa.DisplayName
                Write-Host "Scanning Cral Rules for Search Service Application [$i/$total] {$serviceName}"

                $crawlRules = Get-SPEnterpriseSearchCrawlRule -SearchApplication $ssa

                $j = 1
                $totalCR = $crawlRules.Length
                foreach($crawlRule in $crawlRules)
                {
                    $crPath = $crawlRule.Path
                    Write-Host "    -> Scanning Crawl Rule [$j/$totalCR] {$crPath}"

                    $Script:dscConfigContent += "        SPSearchCrawlRule " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $params.ServiceAppName = $serviceName
                    $params.Path = $crPath
                    $params.Remove("CertificateName")
                    $results = Get-TargetResource @params
                    $results = Repair-Credentials -results $results
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                    $j++
                }
            }
            $i++
        }
        catch
        {
            $Script:ErrorLog += "[Search Crawl Rule]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSearchCrawlerImpactRule()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchCrawlerImpactRule\MSFT_SPSearchCrawlerImpactRule.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object -FilterScript{$_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication"}
    foreach($ssa in $ssas)
    {
        try
        {
            if($null -ne $ssa)
            {
                $impactRules = Get-SPEnterpriseSearchSiteHitRule -SearchService $ssa

                foreach($crawlRule in $impactRules)
                {
                    $Script:dscConfigContent += "        SPSearchCrawlerImpactRule " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $params.ServiceAppName = $ssa.DisplayName
                    $params.Path = $crawlRule.Path
                    $params.Remove("CertificateName")
                    $results = Get-TargetResource @params
                    $results = Repair-Credentials -results $results
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
            }
        }
        catch
        {
            $Script:ErrorLog += "[Search Crawler Impact Rule]" + $ssa.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPOfficeOnlineServerBinding()
{
    $WOPIZone = Get-SPWOPIZone
    $bindings = Get-SPWOPIBinding  -WOPIZone $WOPIZone
    try
    {
        if($null -ne $bindings)
        {
            $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPOfficeOnlineServerBinding\MSFT_SPOfficeOnlineServerBinding.psm1")
            Import-Module $module
            $params = Get-DSCFakeParameters -ModulePath $module

            $Script:dscConfigContent += "        SPOfficeOnlineServerBinding " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
    catch
    {
        $Script:ErrorLog += "[Office Online Server Binding]`r`n"
        $Script:ErrorLog += "$_`r`n`r`n"
    }
}

function Read-SPIrmSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPIrmSettings\MSFT_SPIrmSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $Script:dscConfigContent += "        SPIrmSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params

    $results = Repair-Credentials -results $results

    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPHealthAnalyzerRuleState()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPHealthAnalyzerRuleState\MSFT_SPHealthAnalyzerRuleState.psm1")
    $caWebapp = Get-SPWebApplication -IncludeCentralAdministration `
        | Where-Object -FilterScript{$_.IsAdministrationWebApplication}
    $caWeb = Get-SPWeb($caWebapp.Url)
    $healthRulesList = $caWeb.Lists | Where-Object -FilterScript{$_.BaseTemplate -eq "HealthRules"}

    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    foreach($healthRule in $healthRulesList.Items)
    {
        try
        {
            $params.Name = $healthRule.Title
            $results = Get-TargetResource @params

            if($null -ne $results.Schedule)
            {
                $Script:dscConfigContent += "        SPHealthAnalyzerRuleState " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"

                if($results.Get_Item("Schedule") -eq "On Demand")
                {
                    $results.Schedule = "OnDemandOnly"
                }

                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
            else
            {
                $ruleName = $healthRule.Title
                Write-Warning "Could not extract information for rule {$ruleName}. There may be some missing service applications."
            }
        }
        catch
        {
            $Script:ErrorLog += "[Health Analyzer Rule]" + $healthRule.Title + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPFarmSolution()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarmSolution\MSFT_SPFarmSolution.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $solutions = Get-SPSolution

    foreach($solution in $solutions)
    {
        try
        {
            $Script:dscConfigContent += "        SPFarmSolution " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $solution.Name
            $results = Get-TargetResource @params
            if($results.ContainsKey("ContainsGlobalAssembly"))
            {
                $results.Remove("ContainsGlobalAssembly")
            }
            $filePath = "`$AllNodes.Where{`$Null -ne `$_.SPSolutionPath}.SPSolutionPath+###" + $solution.Name + "###"
            $results["LiteralPath"] = $filePath
            $results = Repair-Credentials -results $results

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentblock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "LiteralPath"
            $currentBlock = $currentBlock.Replace("###", "`"")
            $Script:dscConfigContent += $currentBlock

            $Script:dscConfigContent += "        }`r`n"
        }
        catch
        {
            $Script:ErrorLog += "[Farm Solution]" + $solution.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Save-SPFarmsolution($Path)
{
    Add-ConfigurationDataEntry -Node $env:COMPUTERNAME -Key "SPSolutionPath" -Value $Path -Description "Path where the custom solutions (.wsp) to be installed on the SharePoint Farm are located (local path or Network Share);"
    $solutions = Get-SPSolution
    $farm = Get-SPFarm
    foreach($solution in $solutions)
    {
        try
        {
            $file = $farm.Solutions.Item($solution.Name).SolutionFile
            $filePath = $Path + $solution.Name
            $file.SaveAs($filePath)
        }
        catch
        {
            $Script:ErrorLog += "[Saving Farm Solution]" + $solution.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPFarmAdministrators()
{
    try
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarmAdministrators\MSFT_SPFarmAdministrators.psm1")
        Import-Module $module
        $params = Get-DSCFakeParameters -ModulePath $module
        $params.Remove("MembersToInclude")
        $params.Remove("MembersToExclude")
        $Script:dscConfigContent += "        SPFarmAdministrators " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params
        $results.Name = "SPFarmAdministrators"
        $results = Repair-Credentials -results $results

        $results.Members = Set-SPFarmAdministrators $results.Members
        $results.MembersToInclude = Set-SPFarmAdministrators $results.MembersToInclude
        $results.MembersToExclude = Set-SPFarmAdministrators $results.MembersToExclude

        $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $currentBlock = Set-SPFarmAdministratorsBlock -DSCBlock $currentBlock -ParameterName "Members"
        $currentBlock = Set-SPFarmAdministratorsBlock -DSCBlock $currentBlock -ParameterName "MembersToInclude"
        $currentBlock = Set-SPFarmAdministratorsBlock -DSCBlock $currentBlock -ParameterName "MembersToExclude"
        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
    catch
    {
        $Script:ErrorLog += "[Farm Administrators]`r`n"
        $Script:ErrorLog += "$_`r`n`r`n"
    }
}

function Set-SPFarmAdministratorsBlock($DSCBlock, $ParameterName)
{
    $newLine = $ParameterName + " = @("
    $startPosition = $DSCBlock.IndexOf($ParameterName + " = @")
    if($startPosition -ge 0)
    {
        $endPosition = $DSCBlock.IndexOf("`r`n", $startPosition)
        if($endPosition -ge 0)
        {
            $DSCLine = $DSCBlock.Substring($startPosition, $endPosition - $startPosition)
            $originalLine = $DSCLine
            $DSCLine = $DSCLine.Replace($ParameterName + " = @(","").Replace(");","").Replace(" ","")
            $members = $DSCLine.Split(',')

            foreach($member in $members)
            {
                if($member.StartsWith("`"`$"))
                {
                    $newLine += $member.Replace("`"","") + ", "
                }
                else
                {
                    $newLine += $member + ", "
                }
            }
            if($newLine.EndsWith(", "))
            {
                $newLine = $newLine.Remove($newLine.Length - 2, 2)
            }
            $newLine += ");"
            $DSCBlock = $DSCBlock.Replace($originalLine, $newLine)
        }
    }

    return $DSCBlock
}

function Set-SPFarmAdministrators($members)
{
    $newMemberList = @()
    foreach($member in $members)
    {
        if(!($member.ToUpper() -like "BUILTIN*"))
        {
            $memberUser = Get-Credentials -UserName $member
            if($null -eq $memberUser)
            {
                Save-Credentials -UserName $member
            }
            $accountName = Resolve-Credentials -UserName $member
            $newMemberList += $accountName + ".UserName"
        }
        else
        {
            $newMemberList += $member
        }
    }
    return $newMemberList
}

function Read-SPExcelServiceApp()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPExcelServiceApp\MSFT_SPExcelServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $excelSSAs = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "Excel Services Application Web Service Application"}

    foreach($excelSSA in $excelSSAs)
    {
        try
        {
            if($null -ne $excelSSA)
            {
                $Script:dscConfigContent += "        SPExcelServiceApp " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.Name = $excelSSA.DisplayName
                $results = Get-TargetResource @params
                $privateK = $results.Get_Item("PrivateBytesMax")
                $unusedMax = $results.Get_Item("UnusedObjectAgeMax")

                <# Nik20170106 - Temporary fix while waiting to hear back from Brian F. on how to properly pass these params. #>
                if($results.ContainsKey("TrustedFileLocations"))
                {
                    $results.Remove("TrustedFileLocations")
                }

                if($results.ContainsKey("PrivateBytesMax") -and $privateK -eq "-1")
                {
                    $results.Remove("PrivateBytesMax")
                }

                if($results.ContainsKey("UnusedObjectAgeMax") -and $unusedMax -eq "-1")
                {
                    $results.Remove("UnusedObjectAgeMax")
                }
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Excel Service Application]" + $excelSSA.DisplayName + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

<# Nik20170106 - Read the Designer Settings of either the Site Collection or the Web Application #>
function Read-SPDesignerSettings($receiver)
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDesignerSettings\MSFT_SPDesignerSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $params.Url = $receiver[0]
    $params.SettingsScope = $receiver[1]
    $results = Get-TargetResource @params

    <# Nik20170106 - The logic here differs from other Read functions due to a bug in the Designer Resource that doesn't properly obtains a reference to the Site Collection. #>
    if($null -ne $results)
    {
        $Script:dscConfigContent += "        SPDesignerSettings " + $receiver[1] + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        if($receiver.Length -eq 3)
        {
            $Script:dscConfigContent += "            DependsOn = `"[SP" + $receiver[1].Replace("Collection", "") + "]" + $receiver[2] + "`";`r`n"
        }
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPDatabaseAAG()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDatabaseAAG\MSFT_SPDatabaseAAG.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $databases = Get-SPDatabase
    foreach($database in $databases)
    {
        try
        {
            if($null -ne $database.AvailabilityGroup)
            {
                $Script:dscConfigContent += "        SPDatabaseAAG " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.DatabaseName = $database.Name
                $params.AGName = $database.AvailabilityGroup
                $results = Get-TargetResource @params
                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Database AvailabilityGroup]" + $database.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPWebApplicationAppDomain()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebApplicationAppDomain\MSFT_SPWebApplicationAppDomain.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        $webApplicationAppDomains = Get-SPWebApplicationAppDomain -WebApplication $webApp.Url
        foreach($appDomain in $webApplicationAppDomains)
        {
            $params.WebApplication = $webApp.Url
            $Script:dscConfigContent += "        SPWebApplicationAppDomain " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results

            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

function Read-SPWebAppGeneralSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppGeneralSettings\MSFT_SPWebAppGeneralSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        $params.Url = $webApp.Url
        $Script:dscConfigContent += "        SPWebAppGeneralSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params

        $results = Repair-Credentials -results $results
        if($results.TimeZone -eq -1 -or $null -eq $results.TimeZone)
        {
            $results.Remove("TimeZone")
        }
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPWebAppBlockedFileTypes()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebAppBlockedFileTypes\MSFT_SPWebAppBlockedFileTypes.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        $params.Url = $webApp.Url
        $Script:dscConfigContent += "        SPWebAppBlockedFileTypes " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params

        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPFarmPropertyBag()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarmPropertyBag\MSFT_SPFarmPropertyBag.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $farm = Get-SPFarm
    foreach($key in $farm.Properties.Keys)
    {
        $params.Key = $key
        $Script:dscConfigContent += "        SPFarmPropertyBag " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params

        $results = Repair-Credentials -results $results
        $currentBlock = ""
        if($results.Key -eq "SystemAccountName")
        {
            $accountName = Get-Credentials -UserName $results.Value
            if($null -eq $accountName)
            {
                Save-Credentials -UserName $results.Value
            }
            $results.Value = (Resolve-Credentials -UserName $results.Value) + ".UserName"

            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "Value"
        }
        else
        {
            $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        }
        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPUserProfileServiceAppPermissions()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileServiceAppPermissions\MSFT_SPUserProfileServiceAppPermissions.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $proxies = Get-SPServiceApplicationProxy | Where-Object {$_.TypeName -eq "User Profile Service Application Proxy"}

    foreach($proxy in $proxies)
    {
        try
        {
            $params.ProxyName = $proxy.Name
            $Script:dscConfigContent += "        SPUserProfileServiceAppPermissions " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
        catch
        {
            $Script:ErrorLog += "[User Profile Service Application Permissions]" + $proxy.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPUserProfileSyncConnection()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileSyncConnection\MSFT_SPUserProfileSyncConnection.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $userProfileServiceApps = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "User Profile Service Application"}
    $caURL = (Get-SpWebApplication -IncludeCentralAdministration |
        Where-Object -FilterScript{$_.IsAdministrationWebApplication -eq $true}).Url
    $context = Get-SPServiceContext -Site $caURL
    try
    {
        $userProfileConfigManager  = New-Object -TypeName "Microsoft.Office.Server.UserProfiles.UserProfileConfigManager" `
                                                -ArgumentList $context
        if($null -ne $userProfileConfigManager.ConnectionManager)
        {
            $connections = $userProfileConfigManager.ConnectionManager
            foreach($conn in $connections)
            {
                try
                {
                    $params.Name = $conn.DisplayName
                    $params.ConnectionCredentials = $Global:spFarmAccount
                    $params.UserProfileService = $userProfileServiceApps[0].Name
                    $results = Get-TargetResource @params
                    if($null -ne $results)
                    {
                        if($results.ConnectionType -eq "ActiveDirectoryImport")
                        {
                            $results.ConnectionType = "ActiveDirectory"
                        }
                        $Script:dscConfigContent += "        SPUserProfileSyncConnection " + [System.Guid]::NewGuid().ToString() + "`r`n"
                        $Script:dscConfigContent += "        {`r`n"
                        $results = Repair-Credentials -results $results
                        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                        $Script:dscConfigContent += "        }`r`n"
                    }
                }
                catch
                {
                    $Script:ErrorLog += "[User Profile Sync Connection]" + $conn.DisplayName + "`r`n"
                    $Script:ErrorLog += "$_`r`n`r`n"
                }
            }
        }
    }
    catch
    {

    }
}

function Read-SPUserProfileProperty()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileProperty\MSFT_SPUserProfileProperty.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $caURL = (Get-SpWebApplication -IncludeCentralAdministration | Where-Object -FilterScript{$_.IsAdministrationWebApplication -eq $true}).Url
    $context = Get-SPServiceContext -Site $caURL
    try
    {
        $userProfileConfigManager  = New-Object -TypeName "Microsoft.Office.Server.UserProfiles.UserProfileConfigManager" `
                                                -ArgumentList $context
        $properties = $userProfileConfigManager.GetPropertiesWithSection()
        $properties = $properties | Where-Object{$_.IsSection -eq $false}

        $userProfileServiceApp = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "User Profile Service Application"}

        <# WA - Bug in SPDSC 1.7.0.0 if there is a sync connection, then we need to skip the properties. #>
        if($null -ne $userProfileConfigManager.ConnectionManager.PropertyMapping)
        {
            $i = 1;
            $total = $properties.Length;
            foreach($property in $properties)
            {
                try
                {
                    $params.Name = $property.Name
                    Write-Host "    -> Scanning User Profile Property [$i/$total] {$($property.Name)}"
                    $params.UserProfileService = $userProfileServiceApp[0].DisplayName
                    $Script:dscConfigContent += "        SPUserProfileProperty " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $results = Get-TargetResource @params

                    if($results.Type.ToString().ToLower() -eq "string (single value)")
                    {
                        $results.Type = "String"
                    }
                    elseif($results.Type.ToString().ToLower() -eq "string (multi value)")
                    {
                        $results.Type = "StringMultiValue"
                    }
                    elseif($results.Type.ToString().ToLower() -eq "unique identifier")
                    {
                        $results.Type = "Guid"
                    }
                    elseif($results.Type.ToString().ToLower() -eq "big integer")
                    {
                        $results.Type = "BigInteger"
                    }
                    elseif($results.Type.ToString().ToLower() -eq "date time")
                    {
                        $results.Type = "DateTime"
                    }

                    <# WA - Bug in SPDSC 1.7.0.0 where param returned is named UserProfileServiceAppName instead of
                            just UserProfileService. #>
                    if($null -ne $results.Get_Item("UserProfileServiceAppName"))
                    {
                        $results.Add("UserProfileService", $results.UserProfileServiceAppName)
                        $results.Remove("UserProfileServiceAppName")
                    }

                    $results = Repair-Credentials -results $results
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
                catch
                {
                    $Script:ErrorLog += "[User Profile Property]" + $property.Name + "`r`n"
                    $Script:ErrorLog += "$_`r`n`r`n"
                }
                $i++
            }
        }
    }
    catch {}
}

function Read-SPUserProfileSection()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileSection\MSFT_SPUserProfileSection.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $caURL = (Get-SpWebApplication -IncludeCentralAdministration | Where-Object -FilterScript{$_.IsAdministrationWebApplication -eq $true}).Url
    $context = Get-SPServiceContext -Site $caURL
    try
    {
        $userProfileConfigManager  = New-Object -TypeName "Microsoft.Office.Server.UserProfiles.UserProfileConfigManager" `
                                                -ArgumentList $context
        $properties = $userProfileConfigManager.GetPropertiesWithSection()
        $sections = $properties | Where-Object{$_.IsSection -eq $true}

        $userProfileServiceApp = Get-SPServiceApplication | Where-Object{$_.TypeName -eq "User Profile Service Application"}

        foreach($section in $sections)
        {
            try
            {
                $params.Name = $section.Name
                $params.UserProfileService = $userProfileServiceApp[0].DisplayName
                $Script:dscConfigContent += "        SPUserProfileSection " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params

                $results = Repair-Credentials -results $results
                $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
            catch
            {
                $Script:ErrorLog += "[User Profile Section]" + $section.Name + "`r`n"
                $Script:ErrorLog += "$_`r`n`r`n"
            }
        }
    }
    catch {}
}

function Read-SPBlobCacheSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPBlobCacheSettings\MSFT_SPBlobCacheSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        try
        {
            $alternateUrls = $webApp.AlternateUrl

            if($alternateUrls.Length -ge 1)
            {
                <# WA - Due to Bug in SPDSC 1.7.0.0, we can't have two entries for the same Web Application, but
                        with a different zone. Therefore we are limited to keeping one entry only. #>
                $Script:dscConfigContent += "        SPBlobCacheSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.WebAppUrl = $webApp.Url
                $params.Zone = $alternateUrls[0].Zone
                $results = Get-TargetResource @params
                $results = Repair-Credentials -results $results

                Add-ConfigurationDataEntry -Node "NonNodeData" -Key "BlobCacheLocation" -Value $results.Location -Description "Path where the Blob Cache objects will be stored on the servers;"
                $results.Location = "`$ConfigurationData.NonNodeData.BlobCacheLocation"

                $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "Location"
                $Script:dscConfigContent += $currentBlock
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        catch
        {
            $Script:ErrorLog += "[Blob Cache Settings]" + $webApp.Url + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPSubscriptionSettingsServiceApp()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSubscriptionSettingsServiceApp\MSFT_SPSubscriptionSettingsServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication | Where-Object {$_.TypeName -eq "Microsoft SharePoint Foundation Subscription Settings Service Application"}

    foreach($subSetting in $serviceApps)
    {
        $Script:dscConfigContent += "        SPSubscriptionSettingsServiceApp " + $subSetting.Name.Replace(" ", "") + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $subSetting.Name

        $results = Get-TargetResource @params

        if($null -eq $results.DatabaseName)
        {
            $results.Remove("DatabaseName")
        }

        if($null -eq $results.DatabaseServer)
        {
            $results.Remove("DatabaseServer")
        }

        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAppManagementServiceApp()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppManagementServiceApp\MSFT_SPAppManagementServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication | Where-Object {$_.TypeName -eq "App Management Service Application"}

    foreach($appManagement in $serviceApps)
    {
        $Script:dscConfigContent += "        SPAppManagementServiceApp " + $appManagement.Name.Replace(" ", "") + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $appManagement.Name

        $results = Get-TargetResource @params
        <# WA - Fixes a bug in 1.5.0.0 where the Database Name and Server is not properly returned; #>
        $results.DatabaseName = $appManagement.Databases.Name
        $results.DatabaseServer = $appManagement.Databases.NormalizedDataSource
        $results = Repair-Credentials -results $results

        Add-ConfigurationDataEntry -Node "NonNodeData" -Key "DatabaseServer" -Value $results.DatabaseServer -Description "Name of the Database Server associated with the destination SharePoint Farm;"
        $results.DatabaseServer = "`$ConfigurationData.NonNodeData.DatabaseServer"

        $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "DatabaseServer"
        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAppStoreSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppStoreSettings\MSFT_SPAppStoreSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication

    foreach($webApp in $webApps)
    {
        $Script:dscConfigContent += "        SPAppStoreSettings " + $webApp.Name.Replace(" ", "") + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.WebAppUrl = $webApp.Url
        $results = Get-TargetResource @params
        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAntivirusSettings()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAntivirusSettings\MSFT_SPAntivirusSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $Script:dscConfigContent += "        SPAntivirusSettings AntivirusSettings`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $results = Repair-Credentials -results $results
    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPDistributedCacheService()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDistributedCacheService\MSFT_SPDistributedCacheService.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $params.Name = "DistributedCache"
    $results = Get-TargetResource @params
    if($results.Get_Item("Ensure").ToLower() -eq "present" -and $results.Contains("CacheSizeInMB"))
    {
        $Script:dscConfigContent += "        SPDistributedCacheService " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Repair-Credentials -results $results
        $results.Remove("ServerProvisionOrder")

        $serviceAccount = Get-Credentials -UserName $results.ServiceAccount
        if($null -eq $serviceAccount)
        {
            Save-Credentials -UserName $serviceAccount
        }
        $results.ServiceAccount = (Resolve-Credentials -UserName $results.ServiceAccount) + ".UserName"
        $currentBlock = Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $currentBlock = Convert-DSCStringParamToVariable -DSCBlock $currentBlock -ParameterName "ServiceAccount"
        $Script:dscConfigContent += $currentBlock
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPSessionStateService()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSessionStateService\MSFT_SPSessionStateService.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $svc = Get-SPSessionStateService
    if("" -ne $svc.CatalogName)
    {
        $params.DatabaseName = $svc.CatalogName
        $results = Get-TargetResource @params
        $Script:dscConfigContent += "        SPSessionStateService " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Repair-Credentials -results $results

        <# WA - Bug in the Get-TargetResource in SPDSC 1.7.0.0 not returning the proper set of values #>
        $results.DatabaseName = $svc.CatalogName
        $results.DatabaseServer = $svc.ServerName
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPPasswordChangeSettings
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPPasswordChangeSettings\MSFT_SPPasswordChangeSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $farm = Get-SPFarm

    if($null -ne $farm.PasswordChangeEmailAddress)
    {
        $params.MailAddress = $farm.PasswordChangeEmailAddress
        $results = Get-TargetResource @params
        $Script:dscConfigContent += "        SPPasswordChangeSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPServiceAppSecurity()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPServiceAppSecurity\MSFT_SPServiceAppSecurity.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApplications = Get-SPServiceApplication | Where-Object {$_.GetType().Name -ne "SPUsageApplication" -and $_.GetType().Name -ne "StateServiceApplication"}

    foreach($serviceApp in $serviceApplications)
    {
        try
        {
            $params.ServiceAppName = $serviceApp.Name
            $params.SecurityType = "SharingPermissions"

            $fake = New-CimInstance -ClassName Win32_Process -Property @{Handle=0} -Key Handle -ClientOnly
            $params.Members = $fake
            $params.Remove("MembersToInclude")
            $params.Remove("MembersToExclude")
            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results
            $results.Remove("MembersToInclude")
            $results.Remove("MembersToExclude")

            if($results.Members.Count -gt 0)
            {
                $stringMember = ""
                $foundOne = $false
                foreach($member in $results.Members)
                {
                    $stringMember = Get-SPServiceAppSecurityMembers $member
                    if($null -ne $stringMember)
                    {
                        if(!$foundOne)
                        {
                            $Script:dscConfigContent += "        `$members = @();`r`n"
                            $foundOne = $true
                        }
                        $Script:dscConfigContent += "        `$members += " + $stringMember + ";`r`n"
                    }
                }

                if($foundOne)
                {
                    $Script:dscConfigContent += "        SPServiceAppSecurity " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $results.Members = "`$members"
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
            }

            $params.SecurityType = "Administrators"

            $results = Get-TargetResource @params

            $results = Repair-Credentials -results $results
            $results.Remove("MembersToInclude")
            $results.Remove("MembersToExclude")
            $stringMember = ""

            if($results.Members.Count -gt 0)
            {
                $foundOne = $false
                foreach($member in $results.Members)
                {
                    $stringMember = Get-SPServiceAppSecurityMembers $member
                    if($null -ne $stringMember)
                    {
                        if(!$foundOne)
                        {
                            $Script:dscConfigContent += "        `$members = @();`r`n"
                            $foundOne = $true
                        }
                        $Script:dscConfigContent += "        `$members += " + $stringMember + ";`r`n"
                    }
                }

                if($foundOne)
                {
                    $Script:dscConfigContent += "        SPServiceAppSecurity " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $results.Members = "`$members"
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
            }
        }
        catch
        {
            $Script:ErrorLog += "[Service Application Permissions]" + $serviceApp.Name + "`r`n"
            $Script:ErrorLog += "$_`r`n`r`n"
        }
    }
}

function Read-SPPublishServiceApplication()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPPublishServiceApplication\MSFT_SPPublishServiceApplication.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssas = Get-SPServiceApplication | Where-Object{$_.Shared -eq $true}
    foreach($ssa in $ssas)
    {
        $params.Name = $ssa.DisplayName
        $results = Get-TargetResource @params
        $Script:dscConfigContent += "        SPPublishServiceApplication " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Repair-Credentials -results $results
        $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPRemoteFarmTrust()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPRemoteFarmTrust\MSFT_SPRemoteFarmTrust.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $tips = Get-SPTrustedSecurityTokenIssuer
    foreach($tip in $tips)
    {
        $params.Name = $tip.Id
        $was = Get-SPWebApplication
        foreach($wa in $was)
        {
            $site = Get-SPSite $wa.Url -ErrorAction SilentlyContinue
            if($null -ne $site)
            {
                $params.LocalWebAppUrl = $wa.Url
                $results = Get-TargetResource @params
                if($results.Ensure -eq "Present")
                {
                    $Script:dscConfigContent += "        SPRemoteFarmTrust " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                    $results = Repair-Credentials -results $results
                    $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
                    $Script:dscConfigContent += "        }`r`n"
                }
            }
        }
    }
}

function Read-SPAlternateUrl()
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAlternateUrl\MSFT_SPAlternateUrl.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        $alternateUrls = Get-SPAlternateUrl -WebApplication $webApp

        foreach($alternateUrl in $alternateUrls)
        {
            $Script:dscConfigContent += "        SPAlternateUrl " + [System.Guid]::NewGuid().toString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.WebAppName = $webApp.Name
            $params.Zone = $alternateUrl.UrlZone
            $params.Url = $alternateUrl.IncomingUrl
            $results = Get-TargetResource @params
            $results = Repair-Credentials -results $results
            $Script:dscConfigContent += Get-DSCBlock -UseGetTargetResource -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function sets the settings for the Local Configuration Manager (LCM) component on the server we will be configuring using our resulting DSC Configuration script. The LCM component is the one responsible for orchestrating all DSC configuration related activities and processes on a server. This method specifies settings telling the LCM to not hesitate rebooting the server we are configurating automatically if it requires a reboot (i.e. During the SharePoint Prerequisites installation). Setting this value helps reduce the amount of manual interaction that is required to automate the configuration of our SharePoint farm using our resulting DSC Configuration script. #>
function Set-LCM()
{
    $Script:dscConfigContent += "        LocalConfigurationManager"  + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $Script:dscConfigContent += "            RebootNodeIfNeeded = `$True`r`n"
    $Script:dscConfigContent += "        }`r`n"
}

function Invoke-SQL()
{
    param(
        [Parameter(Mandatory=$true)]
        [string]$Server,
        [Parameter(Mandatory=$true)]
        [string]$dbName,
        [Parameter(Mandatory=$true)]
        [string]$sqlQuery
    )

    $ConnectString="Data Source=${Server}; Integrated Security=SSPI; Initial Catalog=${dbName}"

    $Conn= New-Object System.Data.SqlClient.SQLConnection($ConnectString)
    $Command = New-Object System.Data.SqlClient.SqlCommand($sqlQuery,$Conn)
    $Conn.Open()

    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Command
    $DataSet = New-Object System.Data.DataSet
    $Adapter.Fill($DataSet) | Out-Null

    $Conn.Close()
    $DataSet.Tables
}

<## This method is used to determine if a specific PowerShell cmdlet is available in the current Powershell Session. It is currently used to determine wheter or not the user has access to call the Invoke-SqlCmd cmdlet or if he needs to install the SQL Client coponent first. It simply returns $true if the cmdlet is available to the user, or $false if it is not. #>
function Test-CommandExists()
{
    param ($command)

    $errorActionPreference = "stop"
    try
    {
        if(Get-Command $command)
        {
            return $true
        }
    }
    catch
    {
        return $false
    }
}

function Get-SPReverseDSC()
{
    <## Call into our main function that is responsible for extracting all the information about our SharePoint farm. #>
    Orchestrator

    <## Prompts the user to specify the FOLDER path where the resulting PowerShell DSC Configuration Script will be saved. #>
    $fileName = "SPFarmConfig"
    if($Standalone)
    {
        $fileName = "SPStandalone"
    }
    if($Script:ExtractionModeValue -eq 3)
    {
        $fileName += "-Full"
    }
    elseif($Script:ExtractionModeValue -eq 1)
    {
        $fileName += "-Lite"
    }
    $fileName += ".ps1"
    if($OutputFile -eq "")
    {
        $OutputDSCPath = Read-Host "Please enter the full path of the output folder for DSC Configuration (will be created as necessary)"
    }
    else
    {
            $OutputFile = $OutputFile.Replace("/", "\")
            $fileName = $OutputFile.Split('\')[$OutputFile.Split('\').Length -1]
            $OutputDSCPath = $OutputFile.Remove($OutputFile.LastIndexOf('\') + 1, $OutputFile.Length - ($OutputFile.LastIndexOf('\') + 1))
    }

    <## Ensures the specified output folder path actually exists; if not, tries to create it and throws an exception if we can't. ##>
    while (!(Test-Path -Path $OutputDSCPath -PathType Container -ErrorAction SilentlyContinue))
    {
        try
        {
            Write-Output "Directory `"$OutputDSCPath`" doesn't exist; creating..."
            New-Item -Path $OutputDSCPath -ItemType Directory | Out-Null
            if ($?) {break}
        }
        catch
        {
            Write-Warning "$($_.Exception.Message)"
            Write-Warning "Could not create folder $OutputDSCPath!"
        }
        $OutputDSCPath = Read-Host "Please Enter Output Folder for DSC Configuration (Will be Created as Necessary)"
    }
    <## Ensures the path we specify ends with a Slash, in order to make sure the resulting file path is properly structured. #>
    if(!$OutputDSCPath.EndsWith("\") -and !$OutputDSCPath.EndsWith("/"))
    {
        $OutputDSCPath += "\"
    }

    <# Now that we have acquired the output path, save all custom solutions (.wsp) in that directory; #>
    Save-SPFarmsolution($OutputDSCPath)

    <## Save the content of the resulting DSC Configuration file into a file at the specified path. #>
    $outputDSCFile = $OutputDSCPath + $fileName
    $Script:dscConfigContent | Out-File $outputDSCFile

    <# Add the list of all user accounts detected to the configurationdata #>
    if($Global:AllUsers.Length -gt 0)
    {
        $missingUsers = ""
        foreach($missingUser in $Global:AllUsers)
        {
            $missingUsers += "`"" + $missingUser + "`","
        }
        $missingUsers = "@(" + $missingUsers.Remove($missingUsers.Length-1, 1) + ")"
        Add-ConfigurationDataEntry -Node "NonNodeData" -Key "RequiredUsers" -Value $missingUsers -Description "List of user accounts that were detected that you need to ensure exist in the destination environment;"
    }

    if(!$Azure)
    {
        $outputConfigurationData = $OutputDSCPath + "ConfigurationData.psd1"
        New-ConfigurationDataDocument -Path $outputConfigurationData
    }
    else
    {
      $resGroupName = Read-Host "Destination Resource Group Name"
      $automationAccountName = Read-Host "Destination Automation Account Name"

        $azureDeployScriptPath = $OutputDSCPath + "DeployToAzure.ps1"
        $configurationDataContent = Get-ConfigurationDataContent
        $deployScriptContent = "Login-AzureRMAccount`r`n`$configData = " + $configurationDataContent + "`r`n" + `
            "Import-AzureRmAutomationDscConfiguration -SourcePath (Get-Item '.\" + ($Script:configName + ".ps1") + "').FullName -ResourceGroupName `"" + $resGroupName + "`" -AutomationAccountName `"" + $automationAccountName + "`" -Verbose -Published -Force`r`n"  + `
            "Start-AzureRmAutomationDscCompilationJob -ResourceGroupName `"" + $resGroupName + "`" -AutomationAccountName `"" + $automationAccountName + "`" -ConfigurationName `"" + $Script:configName + "`" -ConfigurationData `$configData"
        $deployScriptContent | Out-File $azureDeployScriptPath
    }

    if($Script:ErrorLog)
    {
        $errorLogPath = $OutputDSCPath + "SharePointDSC.Reverse-Errors.log"
        $Script:ErrorLog | Out-File $errorLogPath
    }

    <## Wait a second, then open our $outputDSCPath in Windows Explorer so we can review the glorious output. ##>
    Start-Sleep 1
    Invoke-Item -Path $OutputDSCPath
}

<## This function defines variables of type Credential for the resulting DSC Configuraton Script. Each variable declared in this method will result in the user being prompted to manually input credentials when executing the resulting script. #>
function Set-ObtainRequiredCredentials()
{
    $credsContent = ""

    foreach($credential in $Global:CredsRepo)
    {
        if(!$credential.ToLower().StartsWith("builtin"))
        {
            if(!$Azure)
            {
                $credsContent += "    " + (Resolve-Credentials $credential) + " = Get-Credential -UserName `"" + $credential + "`" -Message `"Please provide credentials`"`r`n"
            }
            else
            {
                $resolvedName = (Resolve-Credentials $credential)
                $credsContent += "    " + $resolvedName + " = Get-AutomationPSCredential -Name " + ($resolvedName.Replace("$", "")) + "`r`n"
            }
        }
    }
    $credsContent += "`r`n"
    $startPosition = $Script:dscConfigContent.IndexOf("<# Credentials #>") + 19
    $Script:dscConfigContent = $Script:dscConfigContent.Insert($startPosition, $credsContent)
}

Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue
$sharePointSnapin = Get-PSSnapin | Where-Object{$_.Name -eq "Microsoft.SharePoint.PowerShell"}
if($null -ne $sharePointSnapin)
{
    Get-SPReverseDSC
}
else
{
    Write-Host "`r`nE102"  -BackgroundColor Red -ForegroundColor Black -NoNewline
    Write-Host "    - We couldn't detect a SharePoint installation on this machine. Please execute the SharePoint ReverseDSC script on an existing SharePoint server."
}
