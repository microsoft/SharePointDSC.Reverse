# ReverseDSC Orchestrator for SharePoint
This module allows you to extract the current configuration of any given SharePoint 2013 or 2016 (and soon 2019) farm as a PowerShell Desired State Configuration (DSC) .ps1 script along with its associated .psd1 Configuration Data File. With these files you can then recreate an exact copy (down to the SPWeb level) of your SharePoint Farm in another environment (cloud or on-premises).

# How to Use
The ReverseDSC Orchestrator for SharePoint only needs to be put on one server in the farm (recommended to be put on a Web-Front-End).

## Install with internet access
If your machine has internet connectivity, it can be automatically installed using PowerShell 5 and above. This will automatically install the orchestrator and all the modules it depends on. The command to run is:

```PowerShell
Install-Script SharePointDSC.Reverse
```

## Install without internet access
If the server doesn't have internet connectivity, then you will need to run the "Install-Script SharePointDSC.Reverse" command from a computer that has PowerShell version 5 or greater and that had internet connectivity and copy the files manually over the server. Once the command has been run on the machine with internet connectivity, copy the following files to the exact same location on the SharePoint server:
* C:\Program Files\WindowsPowerShell\Modules\SharePointDSC  [Entire Folder]
* C:\Program Files\WindowsPowerShell\Modules\ReverseDSC     [Entire Folder]
* C:\Program Files\WindowsPowerShell\Scripts\SharePointDSC.Reverse.ps1  [If folder doesn't exist on server, create it manually]

## Run
Once you have either run Install-Script SharePointDSC.Reverse or have copied over the files, connect to the SharePoint server using the SharePoint farm account and open a new PowerShell console as an administrator on the server. Browse to the C:\Program Files\WindowsPowerShell\Scripts folder and execute the SharePointDSC.Reverse.ps1 script. Upon running the script, you will be prompted to provide the credentials of an account that has Farm Administrator's privileges. When provided with the proper credentials, the script will automatically extract all the components include in the Extraction mode set and will prompt you to specify a locaton where to save the output.

## Output
As an output, the Orchestrator script will produce a .ps1 file representing the configuration logic for your farm, and a .psd1 file that will be used as Configuration Data to compile the associated MOF file. If you are planning on using ReverseDSC to replicate your SharePoint Farm in another environment, you will need to modify the values in the .psd1 Configuration Data file before compiling your MOF files. Any errors encountered during the extraction will be included in a file named DSC.log inside the specified folder.

# Extraction Modes
The ReverseDSC Orchestrator for SharePoint offers different extraction modes. Please refer to the Extraction Mode page to learn more. https://github.com/Microsoft/SharePointDSC.Reverse/wiki/Extraction-Modes

# Configuration Data
Upon extracting the Desired State Configuration from an existing SharePoint farm, the ReverseDSC Orchestrator for SharePoint will generate a Configuration Data file (.psd1) that will expose all the environment's specific variables for your farm. To replicate your SharePoint Farm into a different environment, all you have to do is ensure the values exposed in that file reflect the values for the destination environment. You can find a complete list of all parameters that are exposed in this Configuration Data File (.psd1) on the following Wiki page https://github.com/Microsoft/SharePointDSC.Reverse/wiki/ConfigurationData


# Parity with SharePointDSC
The following Wiki Page describes the parity between SharePointDSC and ReverseDSC by listing the Resources that are currently covered and being extracted.

https://github.com/Microsoft/SharePointDSC.Reverse/wiki/Parity-with-SharePointDSC

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

# How to Contribute?
Please refer to the follow Wiki page if you are interested in contributing to the project: https://github.com/Microsoft/SharePointDSC.Reverse/wiki/Contribute-to-an-Orchestrator-Script


