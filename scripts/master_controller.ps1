#Creator Rajarshi Rahul Dwivedi
function Monitor-Process {
    param (
        [string]$ScriptPath,
        [int]$TimeoutInHours = 3
    )

    # Start the process and get the initial process information
    $process = Start-Process -FilePath "powershell.exe" -ArgumentList " -File `"$ScriptPath`"" -PassThru

    # Define the timeout duration in seconds
    $timeout = $TimeoutInHours *60*60

    # Start the timer
    $startTime = Get-Date

    # Loop until the process exits or the timeout is reached
    while ($true) {
        Start-Sleep -Seconds 60  # Check every 1 minutes

        # Get the current process state
        $currentProcess = Get-Process -Id $process.Id -ErrorAction SilentlyContinue

        # Check if the process has exited
        if (-not $currentProcess) {
            Write-Output "The script $ScriptPath has completed execution."
            break
        }

        # Calculate the elapsed time
        $elapsedTime = (Get-Date) - $startTime
        if ($elapsedTime.TotalSeconds -ge $timeout) {
            Write-Output "The script $ScriptPath has exceeded the timeout of $TimeoutInHours hours. Terminating the process."

            # Ensure the process still exists before attempting to stop it
            $currentProcess = Get-Process -Id $process.Id -ErrorAction SilentlyContinue
            if ($currentProcess) {
                Stop-Process -Id $process.Id -Force
                Write-Output "Process $ScriptPath terminated."
                get-process conhost | where {$_.'HasExited' -eq $false} | stop-process -Force  
                get-process chrome | where {$_.'HasExited' -eq $false} | stop-process -Force
                $script:Runagain=$true
            } else {
                Write-Output "Process $ScriptPath already exited."
            }
            break
        }
    }

    # Cleanup
    if ($process.HasExited -eq $false) {
        $process.WaitForExit()
    }
}


#start-sleep (3600*2)
$script:Runagain=$false
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$log_path="C:\Automated booking\log.txt"

$stopwatch = [System.Diagnostics.Stopwatch]::new()
$main_path='C:\Automated booking'
cd 'C:\Automated booking'


while($true){
"Script Execution started on $((Get-Date).ToString('dd MMM hh:mm tt'))" | out-file $log_path
$stopwatch.start()



#------------------------------------------------2025-------------------------------------------------------------------------------------------


#-------------------------------------------PMS 2025
Monitor-Process -ScriptPath "$main_path\2025\Run_PMS.ps1" -TimeoutInHours 1
Write-Host "PMS exection completed" -ForegroundColor Red
$stopwatch.Elapsed
write-host "Upload to commitment log will run now for PMS"
Monitor-Process -ScriptPath "$main_path\2025\Scripts\Upload_commitment_logs_PMS.ps1" -TimeoutInHours 1
Write-Host "PMS Commitment Log 2025 Upload completed" -ForegroundColor Green
$stopwatch.Elapsed
try{get-process conhost | where {$_.'HasExited' -eq $false} | stop-process -Force -ErrorAction SilentlyContinue}catch{"No trailing process to stop"}


#-------------------------------------------POS 2025
Monitor-Process -ScriptPath "$main_path\2025\Run_POS.ps1" -TimeoutInHours 3
Write-Host "POS exection completed" -ForegroundColor Red
$stopwatch.Elapsed
write-host "Upload to commitment log will run now for POS"
Monitor-Process -ScriptPath "$main_path\2025\Scripts\Upload_commitment_logs_POS.ps1" -TimeoutInHours 2
Write-Host "POS Commitment Log 2025 Upload completed" -ForegroundColor Green
try{get-process conhost | where {$_.'HasExited' -eq $false} | stop-process -Force -e}catch{"No trailing process to stop"}
$stopwatch.Elapsed




#=================================================================Schedule======================================

"`nScript Execution ended on $((Get-Date).ToString('dd MMM hh:mm tt'))`nTime lapsed $($stopwatch.Elapsed.Hours) Hours and $($stopwatch.Elapsed.Minutes) Minutes" | out-file $log_path -Append
Write-host "Script Execution ended on $((Get-Date).ToString('dd MMM hh:mm tt'))`nTime lapsed $($stopwatch.Elapsed.Hours) Hours and $($stopwatch.Elapsed.Minutes) Minutes" -ForegroundColor Cyan 
$stopwatch.reset()
$stopwatch.Stop()
#.\weblogsemail.bat


if((Get-Date).DayOfWeek -eq "Thursday" -and (Get-Date).Hour -ge 14)
{
"`nExecuting DST changes script this only runs on Thursday last ran at $((Get-Date).ToString('dd MMM hh:mm tt'))"  | out-file $log_path -Append
#.\Scripts\DST_changes_19c.ps1
.\Scripts\DST_changes_PMS.ps1
.\Scripts\DST_changes_POS.ps1
.\Scripts\DST_changes_POS_elit.ps1
.\Scripts\run_this_to_get_latest_chromedriver.ps1
}

.\Scripts\upload_To_sharepoint.ps1 'C:\Automated booking\log.txt' 'log'

if($script:Runagain -eq $false)
{
    while((Get-Date).Hour -lt 5){start-sleep -Seconds (3600);write-host "Script will resume in the morning" -BackgroundColor DarkCyan;"`nNext execution will start at 05:00 AM" | out-file $log_path -Append }
    while((Get-Date).Hour -ge 7 -and (Get-Date).Hour -lt 11){start-sleep -Seconds (3600);write-host "Script will resume in afternoon" -BackgroundColor DarkCyan;"`nNext execution will start at 11:00 AM" | out-file $log_path -Append }
    if((Get-Date).Hour -ge 5 -and (Get-Date).Hour -lt 7){continue;"`nNext execution will start at $((Get-Date).ToString('dd MMM hh:mm tt'))" | out-file $log_path -Append}
    elseif((Get-Date).Hour -ge 11 -and (Get-Date).Hour -lt 12){continue;"`nNext execution will start at $((Get-Date).ToString('dd MMM hh:mm tt'))" | out-file $log_path -Append}
    else{
                "`nNext execution will start at $(((Get-Date).AddHours(3)).ToString('dd MMM hh:mm tt'))" | out-file $log_path -Append
                Write-Host "`nWaiting for 3 hours to resume execution `nLast execution ended at $((Get-Date).ToString('dd MMM hh:mm tt'))`nNext execution should start at $(((Get-Date).AddHours(3)).ToString('dd MMM hh:mm tt'))" -BackgroundColor Red;
                start-sleep -Seconds (1800)
        }
}else{$script:Runagain=$false}

while((Get-Date).DayOfWeek -eq "Saturday" -or (Get-Date).DayOfWeek -eq "Sunday"){start-sleep -Seconds (3600*2);write-host "Script will chill on Saturday and Sunday and resume on Monday" -BackgroundColor DarkCyan }
}