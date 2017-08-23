# ReverseDSC for SharePoint
This module allows you to extract the current configuration of any given SharePoint 2013 or 2016 farm as a PowerShell Desired State Configuration (DSC) .ps1 script.

# Parity with SharePointDSC
The following Wiki Page describes the parity between SharePointDSC and ReverseDSC by listing the Resources that are currently covered and being extracted.

https://github.com/NikCharlebois/SharePointDSC.Reverse/wiki/Parity-with-SharePointDSC

# How does it work?
Every DSC module contains one to many DSC resources. For example, the SharePointDSC module contains resources for SPWebApplication, SPServiceInstance, SPSite, SPUserProfileServiceApplication, etc.). For a DSC resource to be considered valid, it needs to implement three functions at a minimum. These are: 

* Get-TargetResource: Gets the current state of the given resource. For example, for the SPWebApplication resource, this method will return the complete status of any instances of a Web Application in the current SharePoint farm, including information about its application pool, its Host Header, the port it runs on, etc.

* Test-TargetResource: Will compare the current state of the resource against what its desired state is supposed to be. If it detects that there are descrepencies between the desired and current state, it returns $false, otherwise it will return $true. If the ConfigurationMode of the Local Configuration Manager is set to ApplyandMonitor or to ApplyandAutocorrect, then the LCM, by default, will call upon this method every 15 minutes (or at any other interval specified in its settings).

* Set-TargetResource: This method is responsible for bringing an instance of the given DSC Resource in its desired state. This method is called upon when you first call into the Start-DSCConfiguration cmdlet to initiate the configuration of an environment to be in its desired state. If the LCM on the node is configured in "ApplyandAutocorrect" and the Test-DSCResource function detects that the server is not in its desired state, then the Set-TargetResource function will be called upon to try to bring it back into its desired state.

The following diagram shows the links between these three functions. Note that the LCM is the component that keeps the Desired State Configuration information into memory:
![DSC Resources Flow](https://i1.wp.com/nikcharlebois.com/wp-content/uploads/2016/12/LCMProcess.png)

ReverseDSC works by dynamically calling into the Get-TargetResource method of each Resources within a given DSC module. In the case of SharePoint for example, it calls into each DSC Resources (e.g. SPWebApplication, SPServiceInstance, etc.) and extracts information about how each component of the farm is currently configured. It produces a resulting DSC configuration script that represents the exact current state of a farm (down to the Web level). It does not include any content such as Content Types, list and libraries, etc. Think of it as only extracting information that is available within Central Administration.

# Usage
ReverseDSC can be used for many reasons, including:
* Replicate an existing SharePoint environment for troubleshooting;
* Analyzing Best Practice of an environment by reading through the Configuration Script;
* Move and on-premise SharePoint farm onto Azure Infrastructure as a service;
* On-board an existing SharePoint farm onto Desired State Configuration to prevent configuration drifts;
* Migrate a SharePoint 2013 environment to SharePoint 2016;
* Document a SharePoint environment;
* Compare the configuration of two environments, or of the same environments but at two different point in time;
* Create Development standalone machines matching production (merging multiple servers onto a single farm deployment);
* etc.

# Installation
1 - On any SharePoint server within an existing 2013 or 2016 farm, install the SharePointDSC module. If the machine has internet connectivity, this is done by running "Install-Module SharePointDSC", otherwise, the module can be manually download from http://github.com/PowerShell/SharePointDSC and copied into the modules folder (e.g. C:\Program files\WindowsPowerShell\Modules).

2 - On that same SharePoint server, download the latest version of the SharePointDSC.Reverse script from here https://github.com/NikCharlebois/SharePointDSC.Reverse/archive/master.zip and put both files (.ps1 and .psm1) in any directory on the server (they both need to be in the same folder). Recommendation is to create a folder under c:\temp and extract both files under that location;

2a - Since both files have been downloaded from the internet, the recommendation is to turn off the Execution Policy on the server by running "Set-ExecutionPolicy Unrestricted" and by unblocking both files using the "Unblock-File &lt;filename&gt;" command.

3 - Run the SharePointDSC.Reverse.ps1 script in an elevated PowerShell session. Upon validating that all prerequisites have been properly installed on the server, it will prompt you to provide the credentials for any account that has farm administrator rights on the farm. This is required for the script to be able to properly extract the configuration values from the existing environment. When running the script, there are several switches that can be used to speed up the extraction process:

* SkipFeatures - Will skip the extraction of the Features status. By default, ReverseDSC for SharePoint will automatically extract the status of every feature (enabled and disabled) at the Farm level, at the Web Application level, at the Site Collection level as well as at the Web level. Not only does this slows down the entire extraction process, it also tends to generate extremely large resulting DSC configuration scripts. In most scenarios, you will want to use this switch to ensure the extraction results are manageable.

* SkipHealthRules - Will skip the extraction of the Health Analyzer Rules. Since by default all Health Analyzer rules are enabled, this switch can be used for environments where you know for a fact rules were not manually disabled. Otherwise, these will always show as "Present".

* SkipWebs - Will skip the Webs completely. The resulting DSC configuration script will contain everything down to the Site Collection if this switch is used. This can be used to speed up the analysis process, but including webs as part of the resulting script can help provide valuable insights for any given farm.

4  - Once the script has finished its execution, it will prompt you to specify the path to a folder where the resulting DSC configuration script will be stored. If the folder you specify doesn't exist, the script will automatically create it and store the resulting script (e.g. SP-Farm.DSC.ps1) in it.