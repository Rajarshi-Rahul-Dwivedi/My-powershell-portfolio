#Script created by Rajarshi
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
#below funtions are the commands to create UI elements not required to go into details for them
Function Get-Folder($initialDirectory) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.RootFolder = 'MyComputer'
    if ($initialDirectory) { $FolderBrowserDialog.SelectedPath = $initialDirectory }
    [void] $FolderBrowserDialog.ShowDialog()
    return $FolderBrowserDialog.SelectedPath
}
function get-UserInput(){
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Please Enter the Following Parameters"
    $objForm.Size = New-Object System.Drawing.Size(420,340)
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(100,225)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #$OKButton.Add_Click({$Script:userInput=$objTextBox.Text;$objForm.Close()})
    <#$OKButton.Add_Click({
    $pattern=$objTextBox.Text
    $File_extension=$objTextBox2.Text
    $size_gt=$objTextBox3.Text
    $objForm.Close()})#>

    $objForm.Controls.Add($OKButton)

    $CANCELButton = New-Object System.Windows.Forms.Button
    $CANCELButton.Location = New-Object System.Drawing.Size(185,225)
    $CANCELButton.Size = New-Object System.Drawing.Size(90,23)
    $CANCELButton.Text = "CANCEL"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $CANCELButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CANCELButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20)
    $objLabel.Size = New-Object System.Drawing.Size(280,30)
    $objLabel.Text = "Matching Keyword"

    $objLabel2 = New-Object System.Windows.Forms.Label
    $objLabel2.Location = New-Object System.Drawing.Size(10,85)
    $objLabel2.Size = New-Object System.Drawing.Size(280,30)
    $objLabel2.Text = "File Extension"

    $objLabel3 = New-Object System.Windows.Forms.Label
    $objLabel3.Location = New-Object System.Drawing.Size(10,150)
    $objLabel3.Size = New-Object System.Drawing.Size(400,30)
    $objLabel3.Text = "File size should be greater than below value (in MB)"

    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objLabel2)
    $objForm.Controls.Add($objLabel3)

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

    $objTextBox3 = New-Object System.Windows.Forms.TextBox
    $objTextBox3.Location = New-Object System.Drawing.Size(10,180)
    $objTextBox3.Size = New-Object System.Drawing.Size(360,30)
    $objTextBox3.Multiline = $true
    $objForm.Controls.Add($objTextBox3)
    

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})

    $result=$objForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $pattern=$objTextBox.Text
        $File_extension=$objTextBox2.Text
        $size_gt=$objTextBox3.Text

    return $pattern,$File_extension,$size_gt
    }
}
function prompt-date($title)
{
    $form = New-Object Windows.Forms.Form -Property @{
    StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
    Size          = New-Object Drawing.Size 380,400
    Text          = "$title"
    Topmost       = $true
    }

$calendar = New-Object Windows.Forms.MonthCalendar -Property @{
    Location = New-Object System.Drawing.Point(25, 10);
    ShowTodayCircle   = $flase
    MaxSelectionCount = 1
}
$form.Controls.Add($calendar)

$okButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 90, 300
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'OK'
    DialogResult = [Windows.Forms.DialogResult]::OK
}
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object Windows.Forms.Button -Property @{
    Location     = New-Object Drawing.Point 180, 300
    Size         = New-Object Drawing.Size 75, 23
    Text         = 'Cancel'
    DialogResult = [Windows.Forms.DialogResult]::Cancel
}
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$result = $form.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $date = $calendar.SelectionStart
    Write-Host "Date selected: $($date.ToShortDateString())"
    return $date
    }
}
Get-date | out-file "$($PWD.Path)\Search_result.txt" -append

#Start of script , here asking for all the input parameters

Write-Host "Hi This Script will Search for the files based on the provided parameters, Please Select the Parent or root folder to search for `n" -ForegroundColor Yellow

#function call and return value is saved in the variable starting with $ sign

$parent=Get-Folder
Write-Host "Please Input in the dialogue Box :`n`nMatching pattern you want to Search for`nFile Type (Extension of files) `nMinimum file Size `nAnd Range of Date in which file was created" -ForegroundColor Cyan
Write-Host "`nYou can leave one or multiple parameters empty. Script will not consider that parameter while searching`n" -ForegroundColor Yellow
$values=get-UserInput
Write-Host "`nPlease Input Start and End Date between which files were created`n" -ForegroundColor Cyan 
$starting_date=prompt-date -title "Start Date"
$end_date=prompt-date -title "End Date"
#$starting_date=Read-Host "`nPlease enter starting date in format MM/DD/YYYY (eg: 06/17/2001)"
#$end_date=Read-Host "`nPlease enter ending date in format MM/DD/YYYY (eg: 11/29/2022)"
#$starting_date=[datetime]::ParseExact($starting_date, "MM/dd/yyyy", $null)


$values[1]=$values[1] -replace '\.',''
#filter 1 searching for files via keyword and extension $files is an array which stores object returned by applied regex these object contains file properties as name size creation date etc

$files=@(Get-ChildItem -Path "$parent\*$($values[0])*.$($values[1])*" -Recurse | %{Write-Host "Examining file: $_" -ForegroundColor DarkGray; $_} |Where { ! $_.PSIsContainer })
Write-Host "Please wait searching for files`n"
Write-Host "Found $($files.Count) files Please wait applying filters`n"
#filter2 if and and condtion used for every element of $files array to check they fall in the specified rang

$filter_date=@()
    foreach($element in $files)
        { 
            
            if(($element.CreationTime -ge $starting_date) -and ($element.CreationTime -le $end_date))
                {
                    $filter_date+=$element
                }
        }
#filter3 searching for file objects on based on Size here i have diveded the size obtained by 1024*1024 to convert it into mega bytes
$filter_size=@()
$total_size=0
    foreach($element in $filter_date)
        {
            
            $size=($element.Length)/(1024*1024)
            $size=[math]::round($size,4)
            if(!([String]::IsNullOrWhiteSpace($values[2])))
                {
                    if($size -ge $values[2])
                        {
                            $filter_size+=$element
                            $total_size+=$size
                        }
                }else{
                        $filter_size+=$element
                        $total_size+=$size
                     }
        }
$filter_size.fullname
$filter_size.fullname | Out-File "$($PWD.Path)\Search_result.txt" -Append
Write-Host "`nSearch Completed `nFiles Found: $($filter_size.count) `nTotal Size: $total_size MB" -ForegroundColor Green
if($filter_size.Count -eq 0)
{
    Write-host "`n`No files Found matching the specified parameter,Please Retry" -ForegroundColor Red
    pause 
    exit
}
Write-host "`n`npress Y if you want to delete all the found files detail files information can be found in log file $($PWD.Path)\Search_result.txt" -ForegroundColor Yellow
$confirm = Read-Host

#if condition for confirmation to delete the files
#iterationg through each file object and deleting, skipping this process if confirmation is no
if($confirm -eq 'Y')
{
     foreach($element in $filter_size)
     {
        try{Remove-Item -Path $element.FullName -Force -ErrorAction Stop}
        catch
        {Write-Host "Unable to Delete $($element.Name)" -ForegroundColor DarkGray
        continue}
        write-host "Deleted - $($element.Name)" -ForegroundColor darkred
     }
     Write-Host "`nFiles Deleted Task Completed" -ForegroundColor Green
}else{Write-Host "`nNo Files Deleted" -ForegroundColor DarkGray}
pause