#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
$date_today=get-date -Format 'dd-MMM_hh-mm_tt'
Start-Transcript -Path "$PSScriptRoot\Execution_log_$($date_today).log" -Append
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$loc=$PSScriptRoot

function check-download {
    $downloadPath = "$env:USERPROFILE\Downloads"

    while ($true) {
        # Check for active Chrome download in progress
        $incomplete = Get-ChildItem -Path $downloadPath -Filter "*.crdownload"
        if ($incomplete) {
            Start-Sleep -Seconds 2
            continue
        }

        # Get target file with the random suffix
        $file = Get-ChildItem -Path $downloadPath -Filter "Recording*.mp4" | Where-Object {
            $_.Name -match '^Recording\s+.* - [a-z0-9]+\.mp4$'
        }

        if ($file) {
            # Wait for file size to stabilize
            $lastSize = $file.Length
            Start-Sleep -Seconds 3
            $newSize = (Get-Item $file.FullName).Length

            if ($lastSize -eq $newSize) {
                # File is stable, rename it
                if ($file -is [array]) {$file = $file[0]}
                $course_name_current=$course_name -replace '[\\/:*?"<>|]', '-'
                $newName = $file.Name -replace '\s+-\s+[a-z0-9]+(?=\.mp4$)', ''
                $newName=$newName -replace "Recording\s" ,"$course_name_current"
                Rename-Item -Path $file.FullName -NewName $newName
                $targetFolder = "E:\$course_name_current"
                    if (-not (Test-Path $targetFolder)) {
                    New-Item -ItemType Directory -Path $targetFolder | Out-Null
                }
                # Move the file
                Move-Item -Path (Join-Path -Path $file.DirectoryName -ChildPath $newName) -Destination $targetFolder

                Write-Host "File '$newName' is downloaded."
                break
            }
        }

        Start-Sleep -Seconds 2
    }
}



function get-courseName {
    [CmdletBinding()]
    param (
        [string]$str3
    )
    Clear-Host
    Write-Host "================ WELCOME To GEOLOGY CONCEPT VIDEO EXTRACTOR SCRIPT ================" -ForegroundColor Cyan

    Write-Host "Please choose course name`n" -ForegroundColor Cyan
    Write-Host "Press '1'   for     Course 1" -ForegroundColor Cyan
    Write-Host "Press '2'   for     Course 2" -ForegroundColor Cyan
    Write-Host "Press '3'   for     Course 3" -ForegroundColor Cyan
    Write-Host "Press '4'   To      Manually enter the course" -ForegroundColor Cyan
    $input=Read-Host "`n"
    switch ($input)
    {
        1 { return "Course 1" }
        2 { return "Course 2" }
        3 { return "Course 3" }
        4 { return  Read-Host "Please enter the Course name manually"}
        default{Write-Host "Please select a Valid Responce" -ForegroundColor Darkred
            Start-Sleep -Seconds 2.5
            $input=get-menu
            return $input}
    }

}
$course_name=get-courseName

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
$content_menu.FindElementByXPath('(.//a[text()="Live Classes"])[2]').click()
Start-Sleep -Seconds 2
$status_filter=$ChromeDriver.FindElementByXPath('//div[@id="status-filter"]')
$status_filter.FindElementByXPath('(.//label[text()="All"])').click()
$ChromeDriver.FindElementByXPath('//div//button[text()=" Reset"]').click()
Start-Sleep -Seconds 2
$ChromeDriver.FindElementByXPath('//span[text()="Add Filters"]').click()
$ChromeDriver.FindElementByXPath('.//a[text()="Course"]').click()
$course_filter=$ChromeDriver.FindElementByXPath('//div[@id="course-filter"]')
$course_search_bar=$course_filter.FindElementByXPath('//input[@type="text" and @placeholder="Search Course"]')
$course_search_bar.SendKeys($course_name)
Start-Sleep -Seconds 2
$course_search_bar.SendKeys([OpenQA.Selenium.Keys]::DOWN)
$course_search_bar.SendKeys([OpenQA.Selenium.Keys]::RETURN)
$ChromeDriver.FindElementByXPath('//span[text()="Search"]').click()
Start-Sleep -Seconds 4
$total_video_count=$ChromeDriver.FindElementByXPath('//div//b[@id="totalLiveClasses"]')
$total_video_count=[int]($total_video_count.Text)
Write-Host "`n For the course $course_name there are $($total_video_count) Videos`n" -ForegroundColor Yellow
$total_video_count=[math]::Ceiling(($total_video_count/10))
#course and filter selected selecting download buttons from the table

write-host "PRESS ENTER AND IGNORE THIS RED PROMT BELOW IF RUNNING FOR THE INTIAL RUN`n`nIf your execution was aborted in between please click on next button till you are at the required page and press enter to continue downloads" -BackgroundColor Red;pause

for($i=1;$i -le $total_video_count;$i++)
{
    start-sleep -seconds 7
    $Video_table=$ChromeDriver.FindElementByXPath('//table[@id="DataTables_Table_0"]')
    $current_page=($ChromeDriver.FindElementByXPath('//div[@id="DataTables_Table_0_info" and @role="status"]')).text
    Write-Host "`n Current working page is $($current_page)`n" -ForegroundColor Green
    $ChromeDriver.ExecuteScript("window.scrollTo(0, 0);")
    for($k=1;$k -le 10;$k++)
    {
        Start-Sleep -Seconds 2
        $Video_table_row=$Video_table.FindElementByXPath(".//tbody//tr[$k][@role='row']")
        try{$Video_table_row.FindElementByXPath('.//i[text()="download"]').click()}catch{write-host "Skipping $k element not clickable";continue}
        start-sleep -seconds 7
        check-download
    }
    $ChromeDriver.FindElementByXPath('//div//a[text()="Next"]').click()
    Start-Sleep -Seconds 2
}
$ChromeDriver.Quit()
Stop-Transcript
