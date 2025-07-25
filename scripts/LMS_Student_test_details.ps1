#script created by Rajarshi
#Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force
#$mineral_data=ConvertFrom-Csv -Delimiter "," -InputObject (gc "$loc\Mineral_info.csv")


$course_name="IIT-JAM 2025-26 Test Series"


$date_today=get-date -Format 'dd-MMM'
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$loc=$PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms

#windows form
$exception_list=gc "$loc\exception_lis.txt"

function Save-StudentDetails {
    param (
        [Parameter(Mandatory)]
        [ref]$studentDetails,
        [Parameter(Mandatory)]
        [string]$url,
        [Parameter()]
        [string]$studentName,
        [Parameter()]
        [string]$testName,
        [Parameter()]
        [string]$topicName
    )

    # Step 1: Store original window
    $originalWindow = $ChromeDriver.CurrentWindowHandle

    # Step 2: Open new tab
    $ChromeDriver.ExecuteScript("window.open('about:blank','_blank');")

    # Step 3: Switch to new tab
    $windows = $ChromeDriver.WindowHandles
    $newWindow = $windows | Where-Object { $_ -ne $originalWindow }
    $ChromeDriver.SwitchTo().Window($newWindow) | out-null
   
    # Step 4: Navigate to URL
    $ChromeDriver.Navigate().GoToUrl($url)
    start-sleep 2
     try{
    try{$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath(".//div[@class='w-full'][1]")).click()}catch{start-sleep 4;$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath(".//div[@class='w-full'][1]")).click()}
    start-sleep 1
    $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath(".//span[text()='Attempt 1']")).click()
    }catch{"Olny one attempt"}
    $main_element=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//div[@class="mb-5"][1]'))
    
    # Extracting values
    $percentage = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Percentage']/following::span[1]")).Text
    $score      = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Your Score']/following::span[1]")).Text
    try{$rank       = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Rank']/following-sibling::p")).Text}catch{"No Rank";$rank="NA"}
    $totalQs    = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Total questions']/following-sibling::p")).Text
    $correctQs  = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Correct questions']/following-sibling::p")).Text
    $incorrectQs= $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Incorrect questions']/following-sibling::p")).Text
    $unanswered = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Unanswered questions']/following-sibling::p")).Text
    $accuracy   = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Accuracy']/following-sibling::p")).Text
    $timeTaken  = $main_element.FindElement([OpenQA.Selenium.By]::XPath(".//p[text()='Total time taken']/following-sibling::p")).Text
    $testName   = $testName -replace "assignment " , ""
    # Build the object
    $studentDetails.Value += [PSCustomObject]@{
    "Course Name"           = $course_name
    "Topic"                 = $topicName
    "Test/Quiz Name"        = $testName
    "Rank"                  = $rank
    "Percentage"            = $percentage
    "Score"                 = $score
    "Total questions"       = $totalQs
    "Correct questions"     = $correctQs
    "Incorrect questions"   = $incorrectQs
    "Unanswered questions"  = $unanswered
    "Accuracy"              = $accuracy
    "Total time taken"      = $timeTaken
    "Detailed Report Link"  = "GCReport Detailed $($studentName) $($testName).pdf"
    }
    
    # Save as PDF via Chrome DevTools
    $printOptions = [OpenQA.Selenium.PrintOptions]::new()
    #$printOptions.Orientation = "landscape" 
    $pdfData = $ChromeDriver.Print($printOptions)
    $pdfPath="$($loc)\$($student_Reports)\pdf\GCReport Detailed $($studentName) $($testName).pdf"
    [IO.File]::WriteAllBytes($pdfPath, [Convert]::FromBase64String($pdfData.AsBase64EncodedString))
    #Close new tab and switch back
    $ChromeDriver.Close()
    $ChromeDriver.SwitchTo().Window($originalWindow) | out-null
    #$studentDetails.value | Format-Table
    return "Completed for $testName"
}

<#
4. Geology Test Series - UPSC Geo Scientist Prelims
5. GS-Test Series for UPSC-GSI
6. Free Aptitude Test
7. Free Geology Quiz
8. Free GS Quiz
9. IIT-JAM 2025-26 Test Series

#>
#$form_input=get-UserInput
#if($form_input -eq $null){exit}
#else{$course_name=$form_input}

$cleanCourse_name=$course_name -replace '[<>:"/\\|?*\x00-\x1F]', '-'
$student_Reports="Reports - $cleanCourse_name"
mkdir "$($loc)\$($student_Reports)"
mkdir "$($loc)\$($student_Reports)\pdf"
#Setting Up Selenium Web Driver
$workingPath = $loc
if (($env:Path -split ';') -notcontains $workingPath) {$env:Path += "$workingPath"}

Add-Type -Path "$($workingPath)\driver\WebDriver.dll"
Add-Type -Path "$($workingPath)\driver\WebDriver.Support.dll"

$ChromeOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$chromeOptions.AddArgument("--start-minimized")
#$ChromeOptions.AddArgument('start-maximized')
#$ChromeOptions.AddArgument("--disable-background-tab-detection")
#$chromeOptions.AddArgument("--headless")
#$chromeOptions.AddArgument("--disable-cache")
#$chromeOptions.AddArgument("--disable-extensions")
#$chromeOptions.AddArgument("--disable-gpu")


#$ChromeOptions.AcceptInsecureCertificates = $True

#loading profile
#$ChromeOptions.AddArgument("--user-data-dir=$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\AppData\Local\Google\Chrome\User Data")
#$ChromeOptions.AddArgument("--profile-directory=Default")
#$ChromeOptions.addargument('profile-directory=Default')

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($ChromeOptions)
$actions = New-Object OpenQA.Selenium.Interactions.Actions($ChromeDriver)

#login
$ChromeDriver.Navigate().GoToURL("")
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//a[text()="Login"]')).Click()
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//input[@type="email" and @id="le"]')).SendKeys("")
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//input[@type="password" and @id="lp"]')).SendKeys("")
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//button[normalize-space(text())="Login"]')).Click()
#pause
#search for course
$ChromeDriver.Navigate().GoToURL("")
Start-Sleep -Seconds 1
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//label[text()=" Content "]')).Click()
Start-Sleep -Seconds 2
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//a[text()="Courses"]')).Click()
Start-Sleep -Seconds 1
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//input[@type="text" and @id="searchCourse"]')).SendKeys("$course_name")
Start-Sleep -Seconds 2
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//input[@type="text" and @id="searchCourse"]')).SendKeys([OpenQA.Selenium.Keys]::DOWN)
Start-Sleep -Seconds 3
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//input[@type="text" and @id="searchCourse"]')).SendKeys([OpenQA.Selenium.Keys]::RETURN)
Start-Sleep -Seconds 3
#$ChromeDriver.ExecuteScript("window.scrollBy(0,2050)", "")
$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//a[text()=" Learners"]')).Click()
$learners_url=$ChromeDriver.Url
start-sleep 2
$totalLearners_item=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//div[@id="DataTables_Table_0_info" and @role="status"]')).text

#$matches.Clear() 
if ($totalLearners_item -match "Showing 1 to 10 of (\d+) entries") {
    $totalLearners = [int]$matches[1]
    [decimal]$totalLearners = $totalLearners / 10;
}
$break_count=0
$completed_email=@()
$completed_email=@(gc "$($loc)\driver\$($cleanCourse_name) Completed emails.txt")
$studentList =@()
$Serial_number=0
#$ChromeDriver.Navigate().GoToURL("$learners_url")
Write-Host "Total Learners Pages $totalLearners"
if(test-path "$($loc)\driver\last_working Page $($cleanCourse_name).txt"){Write-host "`nLast working page was $(gc "$($loc)\driver\last_working Page $($cleanCourse_name).txt")"}
Write-Host "`nAt Learners Page please select the last stopped page or press enter to continue from page one"
pause
for($j=1;$j -le $totalLearners+1;$j++){
    start-sleep 2
    $currentLearnerPage=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//div[@id="DataTables_Table_0_info" and @role="status"]')).text
    if(!([string]::IsNullOrEmpty($currentLearnerPage))){$currentLearnerPage | set-content "$($loc)\driver\last_working Page $($cleanCourse_name).txt"}
    Write-Host $currentLearnerPage
    for($counter=1;$counter -le 10;$counter++){
        start-sleep 1
        $table_studentpage=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//table[@class="table dataTable no-footer"]'))
        $student_name=($table_studentpage.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$counter]//td[3]"))).text
        $CleanName=$student_name -replace '[\\/:*?"<>|]', ''
        $student_email=($table_studentpage.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$counter]//td[4]"))).text
        $student_email=$student_email -replace "open_in_new" ,""
        if($completed_email -match $student_email){echo "Skipping already completed";continue}
        if(!($exception_list -match $student_email)){"$student_email is not on the Required List";continue}

        $expiry=($table_studentpage.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$counter]//td[7]"))).text -replace "edit" , ""
        if([datetime]$expiry -lt (Get-Date).AddDays(-365) ){"$student_name has old expiry skipping";continue}
            
        ($table_studentpage.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$counter]//td[6]"))).text -match "Progress:\s*(\d+%)\s*[\r\n]+Time Taken:\s*(.+)[\r\n]+"
        $progress=$matches[1]
        $time_Taken=$matches[2]
        $Serial_number++
        $studentList += [PSCustomObject]@{
                      "S No."          = $Serial_number
                      "Student Name"   = $student_name;
                      "Email"          = $student_email;
                      "Progress"       = $progress;
                      "Time Taken"     = $time_Taken;
                      "Complete Report"= "GCReport $($CleanName).xlsx"
                      }

        #open completed reports
        if($progress -eq "0%"){continue}
        $studentList | Format-Table
        if([string]::IsNullOrWhiteSpace($student_name)){continue}
        $table_studentpage.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$counter]//td[6]")).FindElement([OpenQA.Selenium.By]::XPath(".//*[text()='Complete Report']")).click()
        start-sleep 2
        $table_Course_items=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('(//table[@class="table"])[1]'))
        $loop_max=($table_Course_items.Text -split "`n" | Where-Object { $_ -match "assignment" -or $_ -match "label_important"}).Count + 2
        $studentDetails=@()
#--------------------------------
        
        for($i=1;$i -le $loop_max;$i++){
            try{$actions.MoveByOffset(1,1).Click().Perform()}
                catch {
                        break 
                      }
            $actions.MoveByOffset(-1, -1).perform()
            $table_Course_items=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('(//table[@class="table"])[1]'))
            try{$Table_row=$table_Course_items.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$i]"))}catch{"Error Caught for $CleanName";continue}
            
            $Test_name=($table_Course_items.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$i]//td[1]"))).text
            $test_name=$Test_name -replace "^(label_important|assignment_turned_in)\s*", ""
            if($Test_name.Contains("label")){$topic_Name=$Test_name -replace "^label_important\s*", "";continue}
            $Test_status=$null
            $Test_status=$table_Course_items.FindElement([OpenQA.Selenium.By]::XPath(".//tbody//tr[$i]//td[2]"))

            if($Test_status.text -eq "Completed"){
                
                try{
                    $button = $Table_row.FindElement([OpenQA.Selenium.By]::XPath(".//td[4]"))
                    if($button.Text -notmatch "info"){continue}
                    $button.FindElement([OpenQA.Selenium.By]::XPath(".//button[.//i[text()='info']]")).click()
                    start-sleep 2
                    $iframe = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath("//iframe[contains(@src, '/t/reports/assessment/')]"))
                    $StudentReportURL = $iframe.GetAttribute("src")
                    #https://courses.geologyconcepts.com/t/reports/assessment/63490025e4b0125abcb45a2a%3A5ca0b29ee4b0a8f922633c87%3A63c85f32e4b07c3c93f1d0e5?lsb
                    #$ChromeDriver.SwitchTo().Frame($iframe) | out-null
                    #$ChromeDriver.SwitchTo().DefaultContent() | out-null
                    start-sleep -Milliseconds 500
                    $actions.MoveByOffset(1,1).Click().Perform()
                    $actions.MoveByOffset(-1, -1).perform()
                    }catch{continue}
                start-sleep -Milliseconds 300
                Save-StudentDetails ([ref]$studentDetails) $StudentReportURL $CleanName $Test_name $topic_Name
                start-sleep -Milliseconds 300
                
                }
            
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        $completed_email+=$student_email
        try{
        $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath(".//a//span[text()='Back to Learners']")).click()
        }catch{start-sleep 1;$actions.MoveByOffset(1,1).Click().Perform()}
        start-sleep 1
        #dump the student details file
        
        $studentDetails | Format-Table
        if($studentDetails.count -ge 1){
        if(test-path "$($loc)\$($student_Reports)\GCReport $($CleanName).xlsx"){Remove-Item "$($loc)\$($student_Reports)\GCReport $($CleanName).xlsx"}
        $studentDetails | Export-Excel "$($loc)\$($student_Reports)\GCReport $($CleanName).xlsx" -WorksheetName "Student Details" -TableStyle Medium6 -AutoSize -FreezeTopRow
        }
    }
        try {
            $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//a[text()="Next"]')).Click()
            }catch {
                    try{$actions.MoveByOffset(1,1).Click().Perform()}
                        catch {
                                break 
                              }
                    $previous_element=$ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('//a[text()="Previous"]'))
                    if($previous_element.text -ne "Previous"){
                        Write-Host "Scrip Stuck please go back to learners page on $course_name"
                        Start-Sleep -Seconds 300
                        $break_count++
                        If($break_coun -gt 3){break}
                    }
                    else{"Course is completed Now";Break}
                   }

}
$completed_email | Set-Content "$($loc)\driver\$($cleanCourse_name) Completed emails.txt"
if(Test-Path "$($loc)\Student List $($cleanCourse_name).xlsx"){Rename-Item "$($loc)\Student List $($cleanCourse_name).xlsx" -NewName "Student List $($cleanCourse_name) $($completed_email.count).xlsx"}
$studentList | Export-Excel "$($loc)\Student List $($cleanCourse_name).xlsx" -WorksheetName "Student List" -TableStyle Medium6 -AutoSize -FreezeTopRow
$ChromeDriver.Quit()
$ChromeDriver.Dispose()

