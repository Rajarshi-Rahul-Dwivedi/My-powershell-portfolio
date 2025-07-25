#script created by Rajarshi
#Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force
$date_today=get-date -Format 'dd-MMM_hh-mm_tt'
Start-Transcript -Path "$PSScriptRoot\Execution_log_$($date_today).log" -Append
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$loc=$PSScriptRoot


function check-download_kapapa {
    [CmdletBinding()]
    param (
        [string]$course_name
    )
    $downloadPath = "$env:USERPROFILE\Downloads"
    $re_run=$true
    while ($true) {
        # Check for active Chrome download in progress
        $incomplete = Get-ChildItem -Path $downloadPath -Filter "*.crdownload"
        if ($incomplete) {
            Start-Sleep -Seconds 2
            continue
        }

        # Get target file with the random suffix
        $file = Get-ChildItem -Path $downloadPath -Filter "*.zip" | Where-Object {
        $_.Name -match '\.zip$' -and $_.LastWriteTime -gt (Get-Date).AddDays(-1)
        }

        if ($file) {
            # Wait for file size to stabilize
            $lastSize = $file.Length
            Start-Sleep -Seconds 2
            $newSize = (Get-Item $file.FullName).Length

            if ($lastSize -eq $newSize) {
                
                if ($file -is [array]) {
                    write-warning "`n Multiple Zip files Found moving to others folder"
                    try{
                        $file | % { Move-Item -Path $_.FullName -Destination "$loc\Export\others" -ErrorAction Stop}
                        }catch
                            {$file | % {Remove-Item  $_.FullName -Force -ErrorAction SilentlyContinue}}
                    return 0
                    }

                $targetFolder = "$loc\Export\$course_name"
                    if (-not (Test-Path $targetFolder)) {
                    New-Item -ItemType Directory -Path $targetFolder | Out-Null
                }
                # Move the file
                try{Move-Item -Path $file.FullName -Destination $targetFolder -ErrorAction Stop}catch{"$($file.name) already exist skipping";Remove-Item $file.FullName}

                Write-Host "File $($file.Name) is downloaded." -ForegroundColor Green
                break
            }
        }

        Start-Sleep -Seconds 15
        if($re_run -eq $false){Write-Warning "Unable to find downloaded zip file";return 1}
        $re_run=$false
    }
}

#Setting Up Selenium Web Driver
$workingPath = $loc
if (($env:Path -split ';') -notcontains $workingPath) {$env:Path += "$workingPath"}
Add-Type -Path "$($workingPath)\driver\WebDriver.dll"


$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$ChromeOptions.AddArgument('start-maximized')
#$ChromeOptions.AcceptInsecureCertificates = $True

#loading profile
#$ChromeOptions.AddArgument("--user-data-dir=$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\AppData\Local\Google\Chrome\User Data")
#$ChromeOptions.AddArgument("--profile-directory=Default")
#$ChromeOptions.addargument('profile-directory=Default')

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
$actions = [OpenQA.Selenium.Interactions.Actions]::new($ChromeDriver)

#login
$ChromeDriver.Navigate().GoToURL("")
$ChromeDriver.FindElementByXPath('//a[text()="Login"]').Click()
$ChromeDriver.FindElementByXPath('//input[@type="email" and @id="le"]').SendKeys("")
$ChromeDriver.FindElementByXPath('//input[@type="password" and @id="lp"]').SendKeys("")
$ChromeDriver.FindElementByXPath('//button[normalize-space(text())="Login"]').Click()
start-sleep -Seconds 2
#Logic start

$content_menu=$ChromeDriver.FindElementByXPath('//div[.//label[text()=" Content "]]')
$content_menu.FindElementByXPath('.//label[text()=" Content "]').Click()
$content_menu.FindElementByXPath('(.//a[text()="Asset Library"])').click()
Start-Sleep -Seconds 2

$courses_lists=Get-ChildItem "$loc\Courses"
Write-Host "Following Courses list found:`n"
$courses_lists.name | ForEach-Object { Write-Host $_ }
foreach($files in $courses_lists){
    $course_name=$files.name -replace ".txt",""
    
    $test_list=Get-Content $files.FullName
    Write-Host "Working on $course_name with $($test_list.count) Tests/Quiz"
    foreach($test_name in $test_list)
    {
        #logic to download
        Write-Host "`n Downloading $test_name `n" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        if($test_name.Contains(",")){$test_name = $test_name.Split(",")[0]}
        $course_search_bar=$ChromeDriver.FindElementByXPath('//input[@type="text" and @placeholder="Title"]')
        $course_search_bar.clear()
        $course_search_bar.SendKeys($test_name)
        Start-Sleep -Seconds 2
        $course_search_bar.SendKeys([OpenQA.Selenium.Keys]::DOWN)
        $course_search_bar.SendKeys([OpenQA.Selenium.Keys]::RETURN)
        Start-Sleep -Seconds 1
        try{
        $ChromeDriver.FindElementByXPath('//span[text()="Search"]').click()
        Start-Sleep -Seconds 4
        $ChromeDriver.FindElementByXPath('//a//i[text()="cloud_download"]').click()
        }catch{write-warning "Course $test_name not found";continue}
        check-download_kapapa $course_name
    }

}
$ChromeDriver.Quit()
Stop-Transcript

