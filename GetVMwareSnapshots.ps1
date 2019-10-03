# See all Snapshots on VMware Vcenter Machines
# Just set your vCenter IP Address and run this script. In Excecution, this script ask for vCenter Credentials.
# Author: Gilberto Jardim Junior
# Supress Certificate Warning
$WarningPreference = 'SilentlyContinue'
$VIServer = "192.168.201.10"


# Verify if vCenter PowerCLI Modules exists and import it
if (!(Get-module | where {$_.Name -eq "VMware.VimAutomation.Core"})) 
	{
		Import-Module VMware.VimAutomation.Core
	}

Connect-VIServer -Server $VIServer
Get-VM | Get-Snapshot | Select VM,Name

Read-Host -Prompt "Snapshot and VM's above should be verified. If nothing listed, is no snapshots. Press Enter to exit."
