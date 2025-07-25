using namespace System.Net

param($Request, $TriggerMetadata)

# Input bindings
$resourceGroupName = ""
$vmName = ""
$serviceName = ""
$vmScriptPath = ""

Install-Module -Name Az.Compute
Import-Module Az.Compute

# Authenticate using managed identity or service principal
Connect-AzAccount -Identity

Write-Host "Connecting to VM $vmName in Resource Group $resourceGroupName..."

# PowerShell script to check service status and list subfolders
$command = @"
Get-Service -Name '$serviceName' | Select-Object Status;
Get-ChildItem -Path '$vmScriptPath' | Select-Object FullName;
"@

# Run the command on the VM using Azure Run Command
$result = Invoke-AzVMRunCommand -ResourceGroupName $resourceGroupName -VMName $vmName -CommandId 'RunPowerShellScript' -ScriptString $command
$vm_list=gc @("<path to VM list>")

foreach($vm in $vm_list ){
Invoke-AzVMRunCommand -ResourceGroupName $resource_group -Name $vm `
  -CommandId 'RunPowerShellScript' `
  -ScriptString 'stop-Service -Name "service1" -force'
  }
# Parse the result
$output = $result.Value[0].Message

# Return HTTP response
$Response = @{
    status = "success"
    details = $output
}

Write-Host "Command Output: $output"

return $Response | ConvertTo-Json -Depth 10