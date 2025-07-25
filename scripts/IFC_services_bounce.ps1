#Hello, This Script will not function in it's current form please read the discription to update the script parameters or feel free to contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281

<#
1.Stop IFC8 Services, 
    IFC8.NET-SERVICE (MANOLD-OPI-EFT)
 
2.Stop IFC Controller, 
    Opera IFC Controller
 
3.Stop the OPI service in the following order  
    1.Stop the OPI Config service, 
        OPI Config Service
    2.Stop the OPI Service, 
        OPI Service
    3.Stop the OPI Utility Service, 
        OPI Utility Service
 
4.Start in the order:  
    1.OracleServiceWhitbreadOPI
        OracleServiceWHITBREADOPI
 
    2.OracleDB19TNSListner
        OracleOraDB19Home1TNSListener
 
    3.IFC Controller, 
        Opera IFC Controller
4.OPI Utility Service, 
5.OPI Service, 
6.OPI Config service,
 
5.Start the IFC8 Service one by one
#>
start-transcript -Path "C:\temp\Transcript_Ifc_services_bounce..txt" -append
function initiate_stop
{
    param (
        [string]$service_name
    )
    Write-Host "Stopping Service $service_name `n" -ForegroundColor Green
    Stop-Service -DisplayName $service_name
    Start-Sleep -Seconds 5
}


function initiate_start
{
    param (
        [string]$service_name
    )
    Write-Host "Starting Service $service_name `n" -ForegroundColor Green
    Start-Service -DisplayName $service_name
    Start-Sleep -Seconds 3
}


$Ifc_services=Get-Service -DisplayName "IFC8.Net*"
$all_services="Opera IFC Controller","OPI Config Service","OPI Service","OPI Utility Service"
$all_services+=$Ifc_services.DisplayName
"Status of all services $(get-date)" | Out-File "C:\temp\Log_Ifc_services_bounce..txt"
Get-Service -displayname $all_services | Out-File "C:\temp\Log_Ifc_services_bounce..txt" -Append
write-Host "Service status before bounce `n" -ForegroundColor Cyan
Get-Service -displayname $all_services
$db_client="",""



Write-Host "Stopping IFC8 Services `n" -ForegroundColor Cyan
$Ifc_services | %{if($_.status -eq "Running"){Stop-Service $_;Start-Sleep -Seconds 3}}
Start-Sleep -Seconds 5
initiate_stop "Opera IFC Controller"

initiate_stop "OPI Config Service"

initiate_stop "OPI Service"

initiate_stop "OPI Utility Service"


Get-Service -displayname $all_services | %{if($_.status -eq "Running"){"Unable to stop $($_.displayname)" | Out-File "C:\temp\Log_Ifc_services_bounce..txt" -Append}}

Write-Host "Now initiating start of services `n" -ForegroundColor Cyan
Start-Sleep -Seconds 3
Get-Service -DisplayName $db_client[0] | %{if($_.status -eq "Stopped"){Start-Service $_;Start-Sleep -Seconds 10}}
Get-Service -DisplayName $db_client[1] | %{if($_.status -eq "Stopped"){Start-Service $_;Start-Sleep -Seconds 10}}

initiate_start "Opera IFC Controller"

initiate_start "OPI Utility Service"

initiate_start "OPI Service"

initiate_start "OPI Config Service"

$Ifc_services=Get-Service -DisplayName "IFC8.Net*"
$Ifc_services | %{if($_.status -eq "Stopped"){Start-Service $_;Start-Sleep -Seconds 3}}

Get-Service -displayname $all_services | %{if($_.status -eq "Stopped"){"Unable to Start $($_.displayname)" | Out-File "C:\temp\Log_Ifc_services_bounce..txt" -Append}}



write-Host "Service status After bounce `n" -ForegroundColor Cyan
Get-Service -displayname $all_services
Write-Host "Process completed" -ForegroundColor Cyan



stop-transcript 
