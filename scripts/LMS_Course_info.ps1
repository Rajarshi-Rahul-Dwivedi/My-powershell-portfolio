#Hello, This Script will not function in it's current form please read the discription to update the script acordingly or contact me at https://www.linkedin.com/in/rajarshi-dwivedi-abab7a281
#Get-Process -Name chrome -ErrorAction SilentlyContinue | Stop-Process -Force
#$mineral_data=ConvertFrom-Csv -Delimiter "," -InputObject (gc "$loc\Mineral_info.csv")
$date_today=get-date -Format 'dd-MMM'
$download_location="$([System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile))\downloads"
$loc=$PSScriptRoot
Add-Type -AssemblyName System.Windows.Forms

#windows form

function get-UserInput(){
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Please Enter the Following"
    $objForm.Size = New-Object System.Drawing.Size(420,340)
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(100,200)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $objForm.Controls.Add($OKButton)

    $CANCELButton = New-Object System.Windows.Forms.Button
    $CANCELButton.Location = New-Object System.Drawing.Size(200,200)
    $CANCELButton.Size = New-Object System.Drawing.Size(90,23)
    $CANCELButton.Text = "CANCEL"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CANCELButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CANCELButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20)
    $objLabel.Size = New-Object System.Drawing.Size(280,30)
    $objLabel.Text = "Course Name"

    $objLabel2 = New-Object System.Windows.Forms.Label
    $objLabel2.Location = New-Object System.Drawing.Size(10,85)
    $objLabel2.Size = New-Object System.Drawing.Size(280,30)
    $objLabel2.Text = "Test name"

    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objLabel2)


    $objTextBox = New-Object System.Windows.Forms.TextBox
    $objTextBox.Location = New-Object System.Drawing.Size(10,50)
    $objTextBox.Size = New-Object System.Drawing.Size(360,30)
    $objTextBox.Multiline = $true
    $objForm.Controls.Add($objTextBox)

    $objTextBox2 = New-Object System.Windows.Forms.TextBox
    $objTextBox2.Location = New-Object System.Drawing.Size(10,115)
    $objTextBox2.Size = New-Object System.Drawing.Size(360,30)
    $objTextBox2.Multiline = $true
    $objForm.Controls.Add($objTextBox2)


    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})

    $result=$objForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $Course=$objTextBox.Text
        $test=$objTextBox2.Text
      

    return $Course,$test
    }
}
$form_input=get-UserInput
if($form_input -eq $null){exit}
else{$course_name=$form_input[0];$test_name=$form_input[1]}



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

#Logic start

#variable setup
#$course_name="GATE/NET/UPSC CGSE Live Online Classes 2024-25"
#$user="manisha.baskey9@gmail.com"
#$test_name="#15_Weekly Test (Geochemistry)"
#$test_name='2nd Session_Cumulative Test_1 [Min & Cryst_Structural Geology]'
$master_list=@()



#get student list
$content_menu=$ChromeDriver.FindElementByXPath('//div[.//label[text()=" Content "]]')
$content_menu.FindElementByXPath('.//label[text()=" Content "]').Click()
$content_menu.FindElementByXPath('(.//a[text()="Live Tests"])[2]').click()
start-sleep 1
$student_list_search=$ChromeDriver.FindElementByXPath('//input[@type="text" and @placeholder="Search by title"]')
$student_list_search.SendKeys("$test_name")
$student_list_search.SendKeys([OpenQA.Selenium.Keys]::DOWN)
$student_list_search.SendKeys([OpenQA.Selenium.Keys]::RETURN)
$ChromeDriver.FindElementByXPath('//span[text()="Search"]').click()
Start-Sleep 3
$ChromeDriver.FindElementByXPath('//button[@data-tooltip="Results"]').click()
start-sleep 2
$iframe = $ChromeDriver.FindElementByXPath("//iframe[contains(@src, 'results')]")
$ChromeDriver.SwitchTo().Frame($iframe) | out-null
$ChromeDriver.FindElementByXPath('//select').click()
$ChromeDriver.FindElementByXPath('//select//option[@value=500]').click()
start-sleep 2
$table_test=$ChromeDriver.FindElementByXPath('//table[@id="DataTables_Table_0"]')
#$loop_max=($ChromeDriver.FindElementsByXPath('//table[@id="DataTables_Table_0"]//tbody//tr')).count
start-sleep 4
$rows = $table_test.FindElementsByXPath('.//tbody//tr')
Write-Host "Fetching Students list`n" -ForegroundColor Green
start-sleep 5
foreach ($row in $rows) {
    
   
    $td1 = $row.FindElementByXPath('./td[1]').Text  # Rank
    $td4 = $row.FindElementByXPath('./td[4]').Text  # Student name
    Write-Host "$td1 $td4 .." -ForegroundColor cyan
    $td5 = $row.FindElementByXPath('./td[5]').Text  # Email
    $td6 = $row.FindElementByXPath('./td[6]').Text  # Phone
    $td7 = $row.FindElementByXPath('./td[7]').Text  # Maximum marks
    $td8 = $row.FindElementByXPath('./td[8]').Text  # Marks obtained
    $td9 = $row.FindElementByXPath('./td[9]').Text  # Correct
    $td10 = $row.FindElementByXPath('./td[10]').Text  # Incorrect
    $td12 = $row.FindElementByXPath('./td[12]').Text  # Time taken
    
    $master_list += [pscustomobject]@{
        'Rank' = $td1
        'Student name' = $td4
        'Email' = $td5
        'Phone' = $td6
        'Maximum marks' = $td7
        'Marks obtained' = $td8
        'Correct' = $td9
        'Incorrect' = $td10
        'Time taken' = $td12
        'Status'=$null
    }
}

$ChromeDriver.SwitchTo().DefaultContent() | out-null

#search for course
$ChromeDriver.Navigate().GoToURL("https://courses.geologyconcepts.com/s/dashboard")
$ChromeDriver.FindElementByXPath('//label[text()=" Content "]').Click()
$ChromeDriver.FindElementByXPath('//a[text()="Courses"]').Click()
Start-Sleep -Seconds 1
$ChromeDriver.FindElementByXPath('//input[@type="text" and @id="searchCourse"]').SendKeys("$course_name")
Start-Sleep -Seconds 2
$ChromeDriver.FindElementByXPath('//input[@type="text" and @id="searchCourse"]').SendKeys([OpenQA.Selenium.Keys]::DOWN)
Start-Sleep -Seconds 1
$ChromeDriver.FindElementByXPath('//input[@type="text" and @id="searchCourse"]').SendKeys([OpenQA.Selenium.Keys]::RETURN)
Start-Sleep -Seconds 2
#$ChromeDriver.ExecuteScript("window.scrollBy(0,2050)", "")
#$ChromeDriver.FindElementByXPath("//a[@title=""$course_name""]").Click()
$learners_url=$ChromeDriver.Url


foreach ($row in $master_list) {
$ChromeDriver.Navigate().GoToURL("$learners_url")
start-sleep 1
#Search for learner
$user=$row.'Email'
$rank=$row.'Rank'
try{$ChromeDriver.FindElementByXPath('//a[normalize-space(text())="Learners"]').Click()}catch{write-warning "Course Name is incorrect please run the script again";pause;break;}
$ChromeDriver.FindElementByXPath('//input[@type="text" and @data-key="email"]').SendKeys("$user");Start-Sleep -Seconds 1
$ChromeDriver.FindElementByXPath('//input[@type="text" and @data-key="email"]').SendKeys([OpenQA.Selenium.Keys]::RETURN);Start-Sleep -Seconds 1
try{$ChromeDriver.FindElementByXPath('//button[normalize-space(text())="Complete Report"]').Click()}catch{write-warning "Learner $user not visible";$row.'Status'="Learner not enrolled in Course";continue}
#$ChromeDriver.ExecuteScript("window.scrollBy(0,document.body.scrollHeight)", "")

Start-Sleep -Seconds 2
try{$row_test=$ChromeDriver.FindElementByXPath("//table//tr[.//span[text()='$test_name']]")}catch{write-warning "Student $($row.'Student name') have no record of test $test_name";$row.'Status'="Learner's test is not visible";continue}
$second_column_text = $row_test.FindElementByXPath(".//td[2]").Text

if ($second_column_text -eq "completed") {  
    $button = $row_test.FindElementByXPath(".//td[4]")
    $button.FindElementByXPath(".//button[.//i[text()='info']]").click()
    $row.'Status'="Test Completed"
} else {
    write-warning "$($row.'Student Name') Test not completed";
    $row.'Status'="Test not completed"
    continue
}
start-sleep -Seconds 2
$iframe = $ChromeDriver.FindElementByXPath("//iframe[contains(@src, 'courses')]")
#$iframe = $ChromeDriver.FindElementByXPath("//iframe[@onload and @style='width: 100%;height:100%;border:0;']")
$ChromeDriver.SwitchTo().Frame($iframe) | out-null
$table_questionwise=$ChromeDriver.FindElementByXPath('//table[@class="table questionWiseTable"]')
$loop_max=($table_questionwise.Text -split "`n").count -1
$counter=0       

while($counter -lt $loop_max){
    $counter++
    Write-Host "Fetching for $rank $user .. $counter" -ForegroundColor Yellow
    $sNo=($table_questionwise.FindElementByXPath(".//tbody//tr[$counter]//td[1]")).text
    $question_marks=($table_questionwise.FindElementByXPath(".//tbody//tr[$counter]//td[7]")).text
    $correct_ans=($table_questionwise.FindElementByXPath(".//tbody//tr[$counter]//td[9]")).text
    $Your_ans=($table_questionwise.FindElementByXPath(".//tbody//tr[$counter]//td[10]")).text
    $your_marks=($table_questionwise.FindElementByXPath(".//tbody//tr[$counter]//td[11]")).text


    Add-Member -InputObject $row -MemberType NoteProperty -Name "Q$sNo Maximum marks" -Value $question_marks -Force | Out-Null
    Add-Member -InputObject $row -MemberType NoteProperty -Name "Q$sNo Correct answer" -Value $correct_ans -Force | Out-Null
    Add-Member -InputObject $row -MemberType NoteProperty -Name "Q$sNo Your Answer" -Value $Your_ans -Force | Out-Null
    Add-Member -InputObject $row -MemberType NoteProperty -Name "Q$sNo Your Marks" -Value $your_marks -Force | Out-Null
    #$row."Q$sNo Maximum marks" = $question_marks
    #$row."Q$sNo Correct answer" = $correct_ans
   
    }
$ChromeDriver.SwitchTo().DefaultContent() | out-null

}

$ChromeDriver.SwitchTo().DefaultContent() | out-null
$ChromeDriver.Navigate().GoToURL("https://courses.geologyconcepts.com/s/dashboard")

$ChromeDriver.Quit()

if(!(Get-Module -ListAvailable -Name importexcel))
{
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider -Name NuGet
Install-Module ImportExcel -AllowClobber -Force
Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose
}
$course_name=$course_name -replace '[\\/:*?"<>|]', ''
$test_name=$test_name -replace '[\\/:*?"<>|]', ''
if(Test-Path -LiteralPath  "$loc\GC_Report $date_today - $course_name - $test_name.xlsx"){Remove-Item -LiteralPath "$loc\GC_Report $date_today - $course_name - $test_name.xlsx"}
$master_list | Export-Excel -Path "$loc\GC_Report $date_today - $course_name - $test_name.xlsx"  -WorksheetName "GC_report" -TableStyle Medium6 -AutoSize -FreezeTopRow
Write-Host "`nExcel report created in $loc`n" -ForegroundColor Green
pause